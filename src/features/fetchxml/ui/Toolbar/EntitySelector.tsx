/**
 * Entity selector with Publisher/Solution filtering support
 * Adapts UI based on user privileges (Full Filter / Solutions-Only / Publishers-Only / Metadata-Only / No Access modes)
 *
 * The query scope section (#25): collapses to one line ("Entity · View: … · solutions") with New
 * always visible, and remembers the publisher / solution selection per environment.
 */

import { useState, useEffect, useMemo, useCallback } from "react";
import {
	Combobox,
	Option,
	OptionGroup,
	makeStyles,
	mergeClasses,
	Button,
	Spinner,
	Text,
	useId,
	tokens,
	Tooltip,
	InfoLabel,
	useComboboxFilter,
	type ComboboxProps,
} from "@fluentui/react-components";
import { Add20Regular, ChevronDown20Regular, ChevronUp20Regular, LockClosed20Regular } from "@fluentui/react-icons";
import { Info } from "../LeftPane/PropertiesPanel/fields";
import { useAccessMode } from "../../../../shared/hooks/useAccessMode";
import type { EntityScopeMode } from "../../model/displaySettings";
import { isEntityInScope } from "../../state/metadataScope";
import { metadataCache } from "../../state/cache";
import { usePublisherFilter } from "../../../../shared/hooks/usePublisherFilter";
import { useSolutionFilter } from "../../../../shared/hooks/useSolutionFilter";
import { useLazyMetadata } from "../../../../shared/hooks/useLazyMetadata";
import { useRememberedScope } from "../../../../shared/hooks/useRememberedScope";
import { scopeSummaryParts, scopeSummaryText, type RememberedScope } from "../../model/scopeSelection";
import { LoadViewPicker } from "./LoadViewPicker";
import type { EntityMetadata, LoadedViewInfo, Publisher, Solution } from "../../api/pptbClient";
import { debugLog } from "../../../../shared/utils/debug";

const REMEMBERED_SCOPE_INFO = "Your publisher and solution selection is remembered for this environment.";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: "8px",
		padding: "8px 12px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		containerType: "inline-size",
	},
	// Chevron, title or summary, New (#25)
	header: {
		display: "flex",
		alignItems: "center",
		gap: "6px",
		minWidth: 0,
	},
	// The whole area left of "New" toggles the section (not just the small chevron button)
	toggle: {
		display: "flex",
		alignItems: "center",
		gap: "6px",
		flex: 1,
		minWidth: 0,
		minHeight: "24px",
		cursor: "pointer",
		userSelect: "none",
		borderRadius: tokens.borderRadiusMedium,
		":hover": {
			backgroundColor: tokens.colorSubtleBackgroundHover,
		},
		":focus-visible": {
			outline: `2px solid ${tokens.colorStrokeFocus2}`,
			outlineOffset: "1px",
		},
	},
	chevron: {
		flexShrink: 0,
		color: tokens.colorNeutralForeground2,
		display: "inline-flex",
	},
	summary: {
		flex: 1,
		minWidth: 0,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	// Collapsed: hidden, not unmounted, so the pickers keep their state (the loaded view's name)
	body: {
		display: "flex",
		flexDirection: "column",
		gap: "8px",
	},
	bodyCollapsed: {
		display: "none",
	},
	filtersRow: {
		display: "flex",
		gap: "12px",
		alignItems: "flex-end",
		flexWrap: "wrap",
		// Stack vertically when container is narrow (< 520px)
		"@container (max-width: 519px)": {
			flexDirection: "column",
			alignItems: "stretch",
		},
	},
	entityRow: {
		display: "flex",
		alignItems: "flex-end",
		gap: "12px",
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
	// In a narrow pane (< 300px) "New" keeps only its icon (#53); the tooltip and label still say "New query"
	newButton: {
		flexShrink: 0,
		marginLeft: "auto",
		"@container (max-width: 299px)": {
			minWidth: "28px",
			maxWidth: "28px",
			paddingLeft: "0",
			paddingRight: "0",
			justifyContent: "center",
			columnGap: "0",
			"& .fui-Button__icon": {
				marginLeft: "0",
				marginRight: "0",
			},
		},
	},
	newButtonLabel: {
		"@container (max-width: 299px)": {
			display: "none",
		},
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
	disabledLabel: {
		color: tokens.colorNeutralForegroundDisabled,
	},
	// Two-line options: name, then unique name and version / prefix (#25)
	option: {
		display: "flex",
		flexDirection: "column",
		minWidth: 0,
	},
	optionSecondary: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
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
	viewPickerRow: {
		display: "flex",
		gap: "12px",
		alignItems: "flex-end",
	},
});

/** Wider than a narrow field, so most options stay on two lines (#25) */
const LISTBOX_PROPS = { style: { minWidth: 260 } };

/** "Select solutions…", the one selected name, or "N selected" */
function multiselectPlaceholder(names: string[], empty: string): string {
	if (names.length === 0) return empty;
	if (names.length === 1) return names[0];
	return `${names.length} selected`;
}

const solutionMatches = (sol: Solution, query: string) => {
	const q = query.trim().toLowerCase();
	return !q || sol.friendlyname.toLowerCase().includes(q) || (sol.uniquename ?? "").toLowerCase().includes(q);
};

const publisherMatches = (pub: Publisher, query: string) => {
	const q = query.trim().toLowerCase();
	return (
		!q ||
		pub.friendlyname.toLowerCase().includes(q) ||
		(pub.uniquename ?? "").toLowerCase().includes(q) ||
		(pub.customizationprefix ?? "").toLowerCase().includes(q)
	);
};

interface EntitySelectorProps {
	selectedEntity: string | null;
	onEntityChange: (entityLogicalName: string) => void;
	onNewQuery: () => void;
	/** Callback when a saved view should be loaded - provides full view info for execution optimization */
	onViewLoad?: (viewInfo: LoadedViewInfo) => void;
	/** User-preferred entity scope mode (may be clamped to privilege ceiling) */
	entityScopeMode: EntityScopeMode;
	/** Limit entities and attributes to Advanced Find–eligible ones */
	advancedFindOnly: boolean;
	/** The section is collapsed to one line (#25; a persisted setting) */
	collapsed: boolean;
	onCollapsedChange: (collapsed: boolean) => void;
}

export function EntitySelector({
	selectedEntity,
	onEntityChange,
	onNewQuery,
	onViewLoad,
	entityScopeMode,
	advancedFindOnly,
	collapsed,
	onCollapsedChange,
}: EntitySelectorProps) {
	const styles = useStyles();
	const publisherComboId = useId("publisher-combobox");
	const solutionComboId = useId("solution-combobox");
	const entityComboId = useId("entity-combobox");
	const bodyId = useId("query-scope");

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

	// Get full entity metadata for the selected entity (needed for LoadViewPicker). An entity outside
	// the selected solutions (from Parse to Tree, a file, a view) comes from the full entity list (#57).
	const selectedEntityMetadata = useMemo(() => {
		if (!selectedEntity) return null;
		return (
			availableEntities.find((e) => e.LogicalName === selectedEntity) ??
			metadataCache.getAllEntityMetadata()?.find((e) => e.LogicalName === selectedEntity) ??
			allEntities.find((e) => e.LogicalName === selectedEntity) ??
			null
		);
	}, [selectedEntity, availableEntities, allEntities]);

	// The view the Load view picker shows, for the collapsed summary (#25)
	const [selectedViewName, setSelectedViewName] = useState<string | null>(null);

	// Handle view selection from LoadViewPicker
	const handleViewSelect = useCallback(
		(viewInfo: LoadedViewInfo) => {
			if (onViewLoad) {
				onViewLoad(viewInfo);
			}
		},
		[onViewLoad],
	);

	// Handle view clear from LoadViewPicker - resets tree to default entity state
	const handleViewClear = useCallback(() => {
		// Re-set the same entity to reset tree to fresh state while keeping entity selected
		if (selectedEntity) {
			onEntityChange(selectedEntity);
		}
	}, [selectedEntity, onEntityChange]);

	// Publisher multiselect with filtering
	const [publisherQuery, setPublisherQuery] = useState("");

	const filteredPublishers = useMemo(
		() => publisherFilter.publishers.filter((pub) => publisherMatches(pub, publisherQuery)),
		[publisherQuery, publisherFilter.publishers],
	);

	// Debug logging for publisher filtering
	useEffect(() => {
		debugLog("entityFilters", "[EntitySelector] Publisher Filter State:", {
			publisherQuery,
			totalPublishers: publisherFilter.publishers.length,
			filteredPublishers: filteredPublishers.length,
			selectedPublisherIds: publisherFilter.selectedPublisherIds,
		});
	}, [
		publisherQuery,
		publisherFilter.publishers,
		filteredPublishers,
		publisherFilter.selectedPublisherIds,
	]);

	// The publisher / solution selection only filters the entity picker: the query is never
	// cleared, and an entity outside the selection shows by its logical name (#57)
	const onPublisherSelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		publisherFilter.updateSelectedPublishers(data.selectedOptions);
		setPublisherQuery("");
	};

	const publisherPlaceholder = useMemo(
		() =>
			multiselectPlaceholder(
				publisherFilter.publishers
					.filter((pub) => publisherFilter.selectedPublisherIds.includes(pub.publisherid))
					.map((pub) => pub.friendlyname),
				"Select publishers…",
			),
		[publisherFilter.selectedPublisherIds, publisherFilter.publishers],
	);

	// Solution multiselect with filtering and grouping
	const [solutionQuery, setSolutionQuery] = useState("");

	// Get current solutions based on mode
	const currentSolutions =
		effectiveScopeMode === "publisher-solution"
			? publisherFilter.solutions
			: solutionFilter.solutions;

	const currentSelectedSolutionIds =
		effectiveScopeMode === "publisher-solution"
			? publisherFilter.selectedSolutionIds
			: solutionFilter.selectedSolutionIds;

	// Group solutions by managed status
	const { unmanagedSolutions, managedSolutions } = useMemo(() => {
		const unmanaged = currentSolutions.filter((sol) => !sol.ismanaged);
		const managed = currentSolutions.filter((sol) => sol.ismanaged);
		return { unmanagedSolutions: unmanaged, managedSolutions: managed };
	}, [currentSolutions]);

	const onSolutionSelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		if (effectiveScopeMode === "publisher-solution") {
			publisherFilter.updateSelectedSolutions(data.selectedOptions);
		} else if (effectiveScopeMode === "solution-only") {
			solutionFilter.updateSelectedSolutions(data.selectedOptions);
		}
		setSolutionQuery("");
	};

	const selectedSolutionNames = useMemo(
		() =>
			currentSolutions
				.filter((sol) => currentSelectedSolutionIds.includes(sol.solutionid))
				.map((sol) => sol.friendlyname),
		[currentSelectedSolutionIds, currentSolutions],
	);

	// Remember the publisher / solution selection for this environment (#25)
	const { publishersWithSolutions, updateSelectedPublishers, updateSelectedSolutions: updatePublisherSolutions } = publisherFilter;
	const { updateSelectedSolutions: updateOnlySolutions } = solutionFilter;
	const availablePublisherIds = useMemo(() => publishersWithSolutions.map((p) => p.publisher.publisherid), [publishersWithSolutions]);
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
		[effectiveScopeMode, updateSelectedPublishers, updatePublisherSolutions, updateOnlySolutions],
	);
	useRememberedScope({
		mode: accessLoading ? null : effectiveScopeMode,
		ready: effectiveScopeMode === "publisher-solution" ? publisherFilter.publishersLoaded : solutionFilter.solutionsLoaded,
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

	const handleNewQuery = () => {
		if (collapsed) onCollapsedChange(false);
		onNewQuery();
	};

	// The collapsed one-line summary (#25)
	const entityDisplayName = selectedEntity
		? selectedEntityMetadata?.DisplayName?.UserLocalizedLabel?.Label || selectedEntity
		: null;
	const summaryParts = scopeSummaryParts({
		entity: entityDisplayName,
		view: selectedEntityMetadata && onViewLoad ? selectedViewName : null,
		solutions: effectiveScopeMode === "all" ? [] : selectedSolutionNames,
	});

	// Loading state during initial access check
	if (accessLoading) {
		return (
			<div className={styles.container}>
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

	const solutionOption = (sol: Solution) => (
		<Option key={sol.solutionid} value={sol.solutionid} text={sol.friendlyname}>
			<span className={styles.option}>
				<span>{sol.friendlyname}</span>
				<span className={styles.optionSecondary}>
					{[sol.uniquename, sol.version && `v${sol.version}`].filter(Boolean).join(" · ")}
				</span>
			</span>
		</Option>
	);

	/** The solution options in Unmanaged / Managed groups, or why there are none */
	const renderSolutionOptions = (emptyMessage?: string) => {
		if (emptyMessage) return <Option disabled>{emptyMessage}</Option>;
		if (unmanagedSolutions.length === 0 && managedSolutions.length === 0) {
			return <Option disabled>No solutions found</Option>;
		}
		const unmanaged = unmanagedSolutions.filter((sol) => solutionMatches(sol, solutionQuery));
		const managed = managedSolutions.filter((sol) => solutionMatches(sol, solutionQuery));
		if (unmanaged.length === 0 && managed.length === 0) {
			return <Option disabled>{`No solutions match "${solutionQuery.trim()}"`}</Option>;
		}
		return (
			<>
				{unmanaged.length > 0 && <OptionGroup label="Unmanaged">{unmanaged.map(solutionOption)}</OptionGroup>}
				{managed.length > 0 && <OptionGroup label="Managed">{managed.map(solutionOption)}</OptionGroup>}
			</>
		);
	};

	const solutionsCombobox = (disabled: boolean, emptyMessage?: string) => (
		<Combobox
			className={styles.combobox}
			input={{ className: styles.comboboxInput }}
			listbox={LISTBOX_PROPS}
			aria-labelledby={solutionComboId}
			placeholder={multiselectPlaceholder(selectedSolutionNames, "Select solutions…")}
			multiselect
			value={solutionQuery}
			selectedOptions={currentSelectedSolutionIds}
			onOptionSelect={onSolutionSelect}
			onChange={(ev) => setSolutionQuery(ev.target.value)}
			disabled={disabled}
		>
			{renderSolutionOptions(emptyMessage)}
		</Combobox>
	);

	return (
		<div className={styles.container}>
			{/* Header: collapse, title or the one-line summary, New (#25) */}
			<div className={styles.header}>
				<div
					className={styles.toggle}
					role="button"
					tabIndex={0}
					title={collapsed ? "Expand query scope" : "Collapse query scope"}
					onClick={() => onCollapsedChange(!collapsed)}
					onKeyDown={(e) => {
						if (e.key === "Enter" || e.key === " ") {
							e.preventDefault();
							onCollapsedChange(!collapsed);
						}
					}}
					aria-expanded={!collapsed}
					aria-controls={bodyId}
				>
					<span className={styles.chevron}>
						{collapsed ? <ChevronDown20Regular /> : <ChevronUp20Regular />}
					</span>
				{collapsed && summaryParts.length > 0 ? (
					<Text className={styles.summary} title={scopeSummaryText(summaryParts)}>
						{entityDisplayName ? (
							<>
								<b>{summaryParts[0]}</b>
								{summaryParts.slice(1).map((part) => ` · ${part}`)}
							</>
						) : (
							scopeSummaryText(summaryParts)
						)}
					</Text>
				) : (
					<Text weight="semibold" className={styles.summary}>
						Query scope
					</Text>
				)}
				</div>
				<Tooltip content="New query" relationship="label">
					<Button className={styles.newButton} appearance="primary" size="small" icon={<Add20Regular />} onClick={handleNewQuery}>
						<span className={styles.newButtonLabel}>New</span>
					</Button>
				</Tooltip>
			</div>

			<div id={bodyId} className={mergeClasses(styles.body, collapsed && styles.bodyCollapsed)}>
				{/* Full Filter Mode: Publisher + Solution filters */}
				{effectiveScopeMode === "publisher-solution" && (
					<div className={styles.filtersRow}>
						{/* Publishers */}
						<div className={styles.field}>
							<div className={styles.label}>
								<InfoLabel id={publisherComboId} info={<Info text={REMEMBERED_SCOPE_INFO} />}>Publishers</InfoLabel>
							</div>
							<Combobox
								className={styles.combobox}
								input={{ className: styles.comboboxInput }}
								listbox={LISTBOX_PROPS}
								aria-labelledby={publisherComboId}
								placeholder={publisherPlaceholder}
								multiselect
								value={publisherQuery}
								selectedOptions={publisherFilter.selectedPublisherIds}
								onOptionSelect={onPublisherSelect}
								onChange={(ev) => setPublisherQuery(ev.target.value)}
								disabled={publisherFilter.publishersLoading}
							>
								{filteredPublishers.length === 0 ? (
									<Option disabled>
										{publisherQuery.trim() ? `No publishers match "${publisherQuery.trim()}"` : "No publishers found"}
									</Option>
								) : (
									filteredPublishers.map((pub) => (
										<Option key={pub.publisherid} value={pub.publisherid} text={pub.friendlyname}>
											<span className={styles.option}>
												<span>{pub.friendlyname}</span>
												<span className={styles.optionSecondary}>
													{[pub.uniquename, pub.customizationprefix && `prefix ${pub.customizationprefix}_`]
														.filter(Boolean)
														.join(" · ")}
												</span>
											</span>
										</Option>
									))
								)}
							</Combobox>
							{publisherFilter.publishersError && (
								<div className={styles.errorText}>{publisherFilter.publishersError}</div>
							)}
						</div>

						{/* Solutions */}
						<div className={styles.field}>
							<div className={styles.label}>
								<InfoLabel id={solutionComboId} info={<Info text={REMEMBERED_SCOPE_INFO} />}>Solutions</InfoLabel>
							</div>
							{solutionsCombobox(
								publisherFilter.selectedPublisherIds.length === 0,
								publisherFilter.selectedPublisherIds.length === 0 ? "Select a publisher first" : undefined,
							)}
						</div>
					</div>
				)}

				{/* Solution-Only Mode: Solution filter (no publisher picker) */}
				{effectiveScopeMode === "solution-only" && (
					<div className={styles.filtersRow}>
						{/* Publisher - disabled with tooltip */}
						<div className={styles.field}>
							<div className={`${styles.label} ${styles.disabledLabel}`}>
								<InfoLabel id={publisherComboId} info={<Info text={REMEMBERED_SCOPE_INFO} />}>Publishers</InfoLabel>
								<Tooltip content="Publisher filtering requires prvReadPublisher privilege" relationship="description">
									<LockClosed20Regular fontSize={14} />
								</Tooltip>
							</div>
							<Combobox
								className={styles.combobox}
								input={{ className: styles.comboboxInput }}
								aria-labelledby={publisherComboId}
								placeholder="No access to publishers"
								disabled
							>
								<Option>Requires prvReadPublisher privilege</Option>
							</Combobox>
						</div>

						{/* Solutions */}
						<div className={styles.field}>
							<div className={styles.label}>
								<InfoLabel id={solutionComboId} info={<Info text={REMEMBERED_SCOPE_INFO} />}>Solutions</InfoLabel>
							</div>
							{solutionsCombobox(solutionFilter.solutionsLoading, solutionFilter.solutionsLoading ? "Loading…" : undefined)}
							{solutionFilter.solutionsError && (
								<div className={styles.errorText}>{solutionFilter.solutionsError}</div>
							)}
						</div>
					</div>
				)}

				{/* Entity selector (all modes) */}
				<div className={styles.entityRow}>
					<div className={styles.field}>
						<label id={entityComboId} className={styles.label}>
							Entity
						</label>
						<Combobox
							className={styles.combobox}
							input={{ className: styles.comboboxInput }}
							aria-labelledby={entityComboId}
							placeholder={entityLoading ? "Loading entities…" : "Select an entity…"}
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
									{effectiveScopeMode !== "all" ? "Select a solution first" : "No entities found"}
								</Option>
							) : (
								filteredEntityOptions
							)}
						</Combobox>
						{entityError && <div className={styles.errorText}>{entityError}</div>}
					</div>
				</div>

				{/* Load View Picker - only show when entity is selected */}
				{selectedEntityMetadata && onViewLoad && (
					<div className={styles.viewPickerRow}>
						<LoadViewPicker
							selectedEntityMetadata={selectedEntityMetadata}
							onViewSelect={handleViewSelect}
							onViewClear={handleViewClear}
							onSelectedViewChange={setSelectedViewName}
						/>
					</div>
				)}

			</div>
		</div>
	);
}
