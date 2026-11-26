/**
 * Entity selector with Publisher/Solution filtering support
 * Adapts UI based on user privileges (Full Filter / Solutions-Only / Publishers-Only / Metadata-Only / No Access modes)
 *
 * The scope chip (#227) heads it: it shows the publisher / solution scope and opens the pickers, the
 * entity source and Advanced Find only. The selection is remembered per environment (#25).
 */

import { useState, useEffect, useMemo, useCallback, type ReactNode } from "react";
import {
	Combobox,
	Option,
	makeStyles,
	Spinner,
	Text,
	useId,
	tokens,
	useComboboxFilter,
	type ComboboxProps,
} from "@fluentui/react-components";
import { LockClosed20Regular } from "@fluentui/react-icons";
import { useAccessMode } from "../../../../shared/hooks/useAccessMode";
import type { EntityScopeMode } from "../../model/displaySettings";
import { isEntityInScope } from "../../state/metadataScope";
import { usePublisherFilter } from "../../../../shared/hooks/usePublisherFilter";
import { useSolutionFilter } from "../../../../shared/hooks/useSolutionFilter";
import { useLazyMetadata } from "../../../../shared/hooks/useLazyMetadata";
import { useRememberedScope } from "../../../../shared/hooks/useRememberedScope";
import { useScopeTip } from "../../../../shared/hooks/useScopeTip";
import type { RememberedScope } from "../../model/scopeSelection";
import { ScopeChip } from "./ScopeChip";
import type { EntityMetadata } from "../../api/pptbClient";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: "8px",
		padding: "8px 12px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		containerType: "inline-size",
	},
	// The scope chip (#227)
	header: {
		display: "flex",
		alignItems: "center",
		gap: "6px",
		minWidth: 0,
	},
	// The builder pane's own command (Collapse, or Pin in a peek, #234), at the header's end
	headerAction: {
		marginLeft: "auto",
		flexShrink: 0,
	},
	// Fluent's Combobox has a built-in minimum width (~250px); in a narrow left pane it drew over the
	// "New" button (#53). The comboboxes shrink with the pane instead.
	combobox: {
		minWidth: 0,
		width: "100%",
	},
	comboboxInput: {
		minWidth: 0,
		width: "100%",
	},
	field: {
		display: "flex",
		flexDirection: "column",
		gap: "4px",
		flex: 1,
		minWidth: 0,
	},
	label: {
		fontSize: "14px",
		fontWeight: 600,
		color: tokens.colorNeutralForeground2,
		display: "flex",
		alignItems: "center",
		gap: "6px",
	},
	note: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
		lineHeight: tokens.lineHeightBase200,
	},
	loadingContainer: {
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		padding: "24px",
	},
	errorText: {
		color: tokens.colorPaletteRedForeground1,
		fontSize: "12px",
		marginTop: "4px",
	},
	noAccessMessage: {
		padding: "16px",
		color: tokens.colorNeutralForeground3,
		fontSize: "14px",
		textAlign: "center",
	},
});

interface EntitySelectorProps {
	selectedEntity: string | null;
	onEntityChange: (entityLogicalName: string) => void;
	/** User-preferred entity scope mode (may be clamped to privilege ceiling) */
	entityScopeMode: EntityScopeMode;
	/** Change the entity source (#227: the scope chip's popover) */
	onEntityScopeModeChange: (mode: EntityScopeMode) => void;
	/** Limit entities and attributes to Advanced Find–eligible ones */
	advancedFindOnly: boolean;
	onAdvancedFindOnlyChange: (value: boolean) => void;
	/** A command beside the scope chip: the builder pane's Collapse or Pin (#234) */
	headerAction?: ReactNode;
}

export function EntitySelector({
	selectedEntity,
	onEntityChange,
	entityScopeMode,
	onEntityScopeModeChange,
	advancedFindOnly,
	onAdvancedFindOnlyChange,
	headerAction,
}: EntitySelectorProps) {
	const styles = useStyles();
	const action = headerAction ? <div className={styles.headerAction}>{headerAction}</div> : null;
	// While the chip isn't shown yet (checking access, no access), the action still is. The row has
	// the loaded header's shape (an empty slot where the chip goes), so the action isn't remounted
	// when access has been checked: focus on it (Collapse, Pin) stays.
	const actionRow = action ? (
		<div className={styles.header}>
			{null}
			{action}
		</div>
	) : null;
	const entityComboId = useId("entity-combobox");
	const entityHintId = useId("entity-hint");
	const entityErrorId = useId("entity-error");

	// Access mode detection
	const { accessSummary, loading: accessLoading, noAccessMode } = useAccessMode();

	// Effective scope mode: user preference clamped to privilege ceiling
	const effectiveScopeMode = useMemo((): EntityScopeMode => {
		if (accessLoading || noAccessMode || !accessSummary) return "all";
		if (entityScopeMode === "publisher-solution" && accessSummary.fullFilterMode)
			return "publisher-solution";
		if (
			entityScopeMode !== "all" &&
			(accessSummary.fullFilterMode || accessSummary.solutionsOnlyMode)
		)
			return "solution-only";
		return "all";
	}, [entityScopeMode, accessSummary, accessLoading, noAccessMode]);

	// Full filter mode (Publisher → Solution → Entity)
	const publisherFilter = usePublisherFilter();

	// Solutions-only mode (Solution → Entity)
	const solutionFilter = useSolutionFilter();

	// Metadata-only mode (all AF-valid entities) or Publishers-only mode
	const { loadEntities } = useLazyMetadata();
	const [allEntities, setAllEntities] = useState<EntityMetadata[]>([]);
	const [allEntitiesLoading, setAllEntitiesLoading] = useState(false);

	// Load all entities when scope mode is "all" (always fetches everything; AF filter is local)
	useEffect(() => {
		if (effectiveScopeMode !== "all") return;

		setAllEntitiesLoading(true);
		loadEntities()
			.then((entities) => setAllEntities(entities))
			.catch((err) => console.error("Failed to load entities:", err))
			.finally(() => setAllEntitiesLoading(false));
	}, [effectiveScopeMode, loadEntities]);

	// Determine available entities based on effective scope mode.
	// advancedFindOnly is applied as a local filter so toggling it is instant.
	const availableEntities = useMemo(() => {
		let entities: EntityMetadata[] = [];
		if (effectiveScopeMode === "publisher-solution") {
			entities = publisherFilter.entities;
		} else if (effectiveScopeMode === "solution-only") {
			entities = solutionFilter.entities;
		} else {
			entities = allEntities;
		}

		// Apply Advanced Find filter locally in every scope — no re-fetch needed when toggled (#57)
		if (advancedFindOnly) {
			entities = entities.filter((e) => isEntityInScope(e, advancedFindOnly));
		}

		return entities;
	}, [
		effectiveScopeMode,
		advancedFindOnly,
		publisherFilter.entities,
		solutionFilter.entities,
		allEntities,
	]);

	// Get current solutions based on mode
	const currentSolutions =
		effectiveScopeMode === "publisher-solution"
			? publisherFilter.solutions
			: solutionFilter.solutions;

	const currentSelectedSolutionIds =
		effectiveScopeMode === "publisher-solution"
			? publisherFilter.selectedSolutionIds
			: solutionFilter.selectedSolutionIds;

	// The publisher / solution selection only filters the entity picker: the query is never
	// cleared, and an entity outside the selection shows by its logical name (#57)
	const onSolutionsChange = useCallback(
		(ids: string[]) => {
			if (effectiveScopeMode === "publisher-solution")
				publisherFilter.updateSelectedSolutions(ids);
			else if (effectiveScopeMode === "solution-only")
				solutionFilter.updateSelectedSolutions(ids);
		},
		[effectiveScopeMode, publisherFilter, solutionFilter],
	);

	// Remember the publisher / solution selection for this environment (#25)
	const {
		publishersWithSolutions,
		updateSelectedPublishers,
		updateSelectedSolutions: updatePublisherSolutions,
	} = publisherFilter;
	const { updateSelectedSolutions: updateOnlySolutions } = solutionFilter;
	const availablePublisherIds = useMemo(
		() => publishersWithSolutions.map((p) => p.publisher.publisherid),
		[publishersWithSolutions],
	);
	const availableSolutionIds = useCallback(
		(publisherIds: string[]) =>
			effectiveScopeMode === "publisher-solution"
				? publishersWithSolutions
						.filter((p) => publisherIds.includes(p.publisher.publisherid))
						.flatMap((p) => p.solutions.map((s) => s.solutionid))
				: solutionFilter.solutions.map((s) => s.solutionid),
		[effectiveScopeMode, publishersWithSolutions, solutionFilter.solutions],
	);
	const applyRememberedScope = useCallback(
		(scope: RememberedScope) => {
			if (effectiveScopeMode === "publisher-solution") {
				updateSelectedPublishers(scope.publisherIds);
				updatePublisherSolutions(scope.solutionIds);
			} else {
				updateOnlySolutions(scope.solutionIds);
			}
		},
		[
			effectiveScopeMode,
			updateSelectedPublishers,
			updatePublisherSolutions,
			updateOnlySolutions,
		],
	);
	const rememberedScope = useRememberedScope({
		mode: accessLoading ? null : effectiveScopeMode,
		ready:
			effectiveScopeMode === "publisher-solution"
				? publisherFilter.publishersLoaded
				: solutionFilter.solutionsLoaded,
		availablePublisherIds,
		availableSolutionIds,
		publisherIds: publisherFilter.selectedPublisherIds,
		solutionIds: currentSelectedSolutionIds,
		apply: applyRememberedScope,
	});

	// Entity search query with filtering
	const [entityQuery, setEntityQuery] = useState<string>("");

	const entityOptions = useMemo(() => {
		return availableEntities.map((entity) => {
			const displayName = entity.DisplayName?.UserLocalizedLabel?.Label || entity.LogicalName;
			const label = `${displayName} (${entity.LogicalName})`;
			return {
				value: entity.LogicalName,
				children: label,
			};
		});
	}, [availableEntities]);

	// Sync entityQuery with selectedEntity to show formatted display name
	useEffect(() => {
		if (selectedEntity) {
			const selectedEntityMetadata = availableEntities.find(
				(e) => e.LogicalName === selectedEntity,
			);
			if (selectedEntityMetadata) {
				const displayName =
					selectedEntityMetadata.DisplayName?.UserLocalizedLabel?.Label ||
					selectedEntityMetadata.LogicalName;
				setEntityQuery(`${displayName} (${selectedEntityMetadata.LogicalName})`);
			} else {
				// Entity not in available list yet, just show logical name
				setEntityQuery(selectedEntity);
			}
		} else {
			setEntityQuery("");
		}
	}, [selectedEntity, availableEntities]);

	const filteredEntityOptions = useComboboxFilter(entityQuery, entityOptions, {
		noOptionsMessage: "No entities found",
	});

	const onEntitySelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		if (data.optionValue) {
			onEntityChange(data.optionValue);
			setEntityQuery(data.optionText ?? "");
		} else {
			setEntityQuery("");
		}
	};

	// The privilege ceiling for the entity source (#227: chosen in the scope chip, no longer in Settings)
	const canUsePublisherSolution = !accessSummary || accessSummary.fullFilterMode;
	const canUseSolutionOnly =
		!accessSummary || accessSummary.fullFilterMode || accessSummary.solutionsOnlyMode;

	// First-run tip (#227): no scope yet, after any remembered one was restored, once per environment
	const scopeTip = useScopeTip();
	const listsReady =
		effectiveScopeMode === "publisher-solution"
			? publisherFilter.publishersLoaded
			: effectiveScopeMode === "solution-only"
				? solutionFilter.solutionsLoaded
				: true;
	const hasScope = effectiveScopeMode === "all" || currentSelectedSolutionIds.length > 0;
	const showScopeTip =
		scopeTip.ready &&
		!scopeTip.dismissed &&
		rememberedScope.restored &&
		listsReady &&
		!hasScope;

	// Loading state during initial access check
	if (accessLoading) {
		return (
			<div className={styles.container}>
				{actionRow}
				<div className={styles.loadingContainer}>
					<Spinner size="small" label="Checking access permissions…" />
				</div>
			</div>
		);
	}

	// No access mode
	if (noAccessMode) {
		return (
			<div className={styles.container}>
				{actionRow}
				<div className={styles.noAccessMessage}>
					<LockClosed20Regular style={{ marginBottom: "8px" }} />
					<div>You don't have permission to read metadata (prvReadCustomization).</div>
					<div style={{ fontSize: "12px", marginTop: "4px" }}>
						Contact your system administrator to request access.
					</div>
				</div>
			</div>
		);
	}

	// Entity loading state
	const entityLoading =
		effectiveScopeMode === "publisher-solution"
			? publisherFilter.entitiesLoading
			: effectiveScopeMode === "solution-only"
				? solutionFilter.entitiesLoading
				: allEntitiesLoading;

	const entityError =
		effectiveScopeMode === "publisher-solution"
			? publisherFilter.entitiesError
			: effectiveScopeMode === "solution-only"
				? solutionFilter.entitiesError
				: null;

	const entityPlaceholder = entityLoading
		? "Loading entities…"
		: !hasScope
			? "Choose a scope to list entities"
			: "Select an entity…";

	return (
		<div className={styles.container}>
			{/* The scope chip (#227); New is in the app bar (#278) */}
			<div className={styles.header}>
				<ScopeChip
					mode={effectiveScopeMode}
					onModeChange={onEntityScopeModeChange}
					canUsePublisherSolution={canUsePublisherSolution}
					canUseSolutionOnly={canUseSolutionOnly}
					publishers={publisherFilter.publishers}
					selectedPublisherIds={publisherFilter.selectedPublisherIds}
					publishersLoading={publisherFilter.publishersLoading}
					publishersError={publisherFilter.publishersError}
					onPublishersChange={publisherFilter.updateSelectedPublishers}
					solutions={currentSolutions}
					selectedSolutionIds={currentSelectedSolutionIds}
					solutionsLoading={
						effectiveScopeMode === "solution-only" && solutionFilter.solutionsLoading
					}
					solutionsError={
						effectiveScopeMode === "solution-only"
							? solutionFilter.solutionsError
							: null
					}
					onSolutionsChange={onSolutionsChange}
					advancedFindOnly={advancedFindOnly}
					onAdvancedFindOnlyChange={onAdvancedFindOnlyChange}
					showTip={showScopeTip}
					onTipDismiss={scopeTip.dismiss}
				/>
				{action}
			</div>

			{/* Entity selector (all modes) */}
			<div className={styles.field}>
				<label id={entityComboId} className={styles.label}>
					Entity
				</label>
				<Combobox
					className={styles.combobox}
					input={{ className: styles.comboboxInput }}
					aria-labelledby={entityComboId}
					aria-describedby={
						entityError ? `${entityHintId} ${entityErrorId}` : entityHintId
					}
					placeholder={entityPlaceholder}
					value={entityQuery}
					onOptionSelect={onEntitySelect}
					onChange={(ev) => setEntityQuery(ev.target.value)}
					clearable
					disabled={entityLoading || availableEntities.length === 0}
				>
					{entityLoading ? (
						<Option disabled>Loading…</Option>
					) : availableEntities.length === 0 ? (
						<Option disabled>
							{!hasScope ? "Choose a scope to list entities" : "No entities found"}
						</Option>
					) : (
						filteredEntityOptions
					)}
				</Combobox>
				<Text id={entityHintId} className={styles.note}>
					{/* While the picker can't be used, the hint says why (#279) */}
					{entityLoading
						? "Loading entities…"
						: availableEntities.length > 0
							? `${availableEntities.length.toLocaleString()} ${availableEntities.length === 1 ? "entity" : "entities"} in scope`
							: !hasScope
								? "Choose a scope to list entities"
								: // A failed load says so below
									entityError
									? ""
									: "No entities in this scope"}
				</Text>
				{entityError && (
					<div id={entityErrorId} className={styles.errorText}>
						{entityError}
					</div>
				)}
			</div>
		</div>
	);
}
