/**
 * Entity selector with Publisher/Solution filtering support
 * Adapts UI based on user privileges (Full Filter / Solutions-Only / Publishers-Only / Metadata-Only / No Access modes)
 */

import { useState, useEffect, useMemo, useCallback } from "react";
import {
	Combobox,
	Option,
	OptionGroup,
	makeStyles,
	Button,
	Spinner,
	useId,
	tokens,
	Tooltip,
	useComboboxFilter,
	type ComboboxProps,
} from "@fluentui/react-components";
import { Add20Regular, LockClosed20Regular } from "@fluentui/react-icons";
import { useAccessMode } from "../../../../shared/hooks/useAccessMode";
import type { EntityScopeMode } from "../../model/displaySettings";
import { isEntityInScope } from "../../state/metadataScope";
import { metadataCache } from "../../state/cache";
import { usePublisherFilter } from "../../../../shared/hooks/usePublisherFilter";
import { useSolutionFilter } from "../../../../shared/hooks/useSolutionFilter";
import { useLazyMetadata } from "../../../../shared/hooks/useLazyMetadata";
import { LoadViewPicker } from "./LoadViewPicker";
import type { EntityMetadata, LoadedViewInfo } from "../../api/pptbClient";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: "8px",
		padding: "8px 12px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		containerType: "inline-size",
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
		"@container (max-width: 299px)": {
			minWidth: "32px",
			maxWidth: "32px",
			paddingLeft: "5px",
			paddingRight: "5px",
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
}

export function EntitySelector({
	selectedEntity,
	onEntityChange,
	onNewQuery,
	onViewLoad,
	entityScopeMode,
	advancedFindOnly,
}: EntitySelectorProps) {
	const styles = useStyles();
	const publisherComboId = useId("publisher-combobox");
	const solutionComboId = useId("solution-combobox");
	const entityComboId = useId("entity-combobox");

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

	const publisherOptions = useMemo(() => {
		return publisherFilter.publishers.map((pub) => ({
			value: pub.publisherid,
			children: pub.friendlyname,
		}));
	}, [publisherFilter.publishers]);

	// Manual filtering for publishers
	const filteredPublisherOptions = useMemo(() => {
		if (!publisherQuery) return publisherOptions;
		const lowerQuery = publisherQuery.toLowerCase();
		return publisherOptions.filter(
			(opt) => typeof opt.children === "string" && opt.children.toLowerCase().includes(lowerQuery),
		);
	}, [publisherQuery, publisherOptions]);

	// Debug logging for publisher filtering
	useEffect(() => {
		console.log("[EntitySelector] Publisher Filter State:", {
			publisherQuery,
			totalPublishers: publisherOptions.length,
			filteredPublishers: filteredPublisherOptions.length,
			selectedPublisherIds: publisherFilter.selectedPublisherIds,
		});
	}, [
		publisherQuery,
		publisherOptions,
		filteredPublisherOptions,
		publisherFilter.selectedPublisherIds,
	]);

	// The publisher / solution selection only filters the entity picker: the query is never
	// cleared, and an entity outside the selection shows by its logical name (#57)
	const onPublisherSelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		publisherFilter.updateSelectedPublishers(data.selectedOptions);
		setPublisherQuery("");
	};

	// Compute publisher placeholder dynamically
	const publisherPlaceholder = useMemo(() => {
		if (publisherFilter.selectedPublisherIds.length === 0) {
			return "Select publishers...";
		}
		const selectedNames = publisherFilter.publishers
			.filter((pub) => publisherFilter.selectedPublisherIds.includes(pub.publisherid))
			.map((pub) => pub.friendlyname);
		return `${selectedNames.length} selected`;
	}, [publisherFilter.selectedPublisherIds, publisherFilter.publishers]);

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

	// Compute solution placeholder dynamically
	const solutionPlaceholder = useMemo(() => {
		if (currentSelectedSolutionIds.length === 0) {
			return "Select solutions...";
		}
		const selectedNames = currentSolutions
			.filter((sol) => currentSelectedSolutionIds.includes(sol.solutionid))
			.map((sol) => sol.friendlyname);
		return `${selectedNames.length} selected`;
	}, [currentSelectedSolutionIds, currentSolutions]);

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
		noOptionsMessage: "No entities match your search.",
	});

	const onEntitySelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		if (data.optionValue) {
			onEntityChange(data.optionValue);
			setEntityQuery(data.optionText ?? "");
		} else {
			setEntityQuery("");
		}
	};

	// Loading state during initial access check
	if (accessLoading) {
		return (
			<div className={styles.container}>
				<div className={styles.loadingContainer}>
					<Spinner size="small" label="Checking access permissions..." />
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

	return (
		<div className={styles.container}>
			{/* Full Filter Mode: Publisher + Solution filters */}
			{effectiveScopeMode === "publisher-solution" && (
				<div className={styles.filtersRow}>
					{/* Publishers */}
					<div className={styles.field}>
						<label id={publisherComboId} className={styles.label}>
							Publishers
						</label>
						<Combobox
							className={styles.combobox}
							input={{ className: styles.comboboxInput }}
							aria-labelledby={publisherComboId}
							placeholder={publisherPlaceholder}
							multiselect
							value={publisherQuery}
							selectedOptions={publisherFilter.selectedPublisherIds}
							onOptionSelect={onPublisherSelect}
							onChange={(ev) => setPublisherQuery(ev.target.value)}
							disabled={publisherFilter.publishersLoading}
						>
							{filteredPublisherOptions.length === 0 ? (
								<Option>No publishers match your search</Option>
							) : (
								filteredPublisherOptions.map((opt) => (
									<Option key={opt.value} value={opt.value}>
										{opt.children}
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
						<label id={solutionComboId} className={styles.label}>
							Solutions
						</label>
						<Combobox
							className={styles.combobox}
							input={{ className: styles.comboboxInput }}
							aria-labelledby={solutionComboId}
							placeholder={solutionPlaceholder}
							multiselect
							value={solutionQuery}
							selectedOptions={currentSelectedSolutionIds}
							onOptionSelect={onSolutionSelect}
							onChange={(ev) => setSolutionQuery(ev.target.value)}
							disabled={publisherFilter.selectedPublisherIds.length === 0}
						>
							{publisherFilter.selectedPublisherIds.length === 0 ? (
								<Option>Select publishers first</Option>
							) : unmanagedSolutions.length === 0 && managedSolutions.length === 0 ? (
								<Option>No solutions found</Option>
							) : (
								<>
									{unmanagedSolutions.length > 0 && (
										<OptionGroup label="Unmanaged">
											{unmanagedSolutions
												.filter((sol) =>
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase()),
												)
												.map((sol) => (
													<Option
														key={sol.solutionid}
														value={sol.solutionid}
														text={sol.friendlyname}
													>
														{sol.friendlyname}
													</Option>
												))}
										</OptionGroup>
									)}
									{managedSolutions.length > 0 && (
										<OptionGroup label="Managed">
											{managedSolutions
												.filter((sol) =>
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase()),
												)
												.map((sol) => (
													<Option
														key={sol.solutionid}
														value={sol.solutionid}
														text={sol.friendlyname}
													>
														{sol.friendlyname}
													</Option>
												))}
										</OptionGroup>
									)}
								</>
							)}
						</Combobox>
					</div>
				</div>
			)}

			{/* Solution-Only Mode: Solution filter (no publisher picker) */}
			{effectiveScopeMode === "solution-only" && (
				<div className={styles.filtersRow}>
					{/* Publisher - disabled with tooltip */}
					<div className={styles.field}>
						<Tooltip
							content="Publisher filtering requires prvReadPublisher privilege"
							relationship="description"
						>
							<label id={publisherComboId} className={`${styles.label} ${styles.disabledLabel}`}>
								Publishers <LockClosed20Regular fontSize={14} />
							</label>
						</Tooltip>
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
						<label id={solutionComboId} className={styles.label}>
							Solutions
						</label>
						<Combobox
							className={styles.combobox}
							input={{ className: styles.comboboxInput }}
							aria-labelledby={solutionComboId}
							placeholder={solutionPlaceholder}
							multiselect
							value={solutionQuery}
							selectedOptions={currentSelectedSolutionIds}
							onOptionSelect={onSolutionSelect}
							onChange={(ev) => setSolutionQuery(ev.target.value)}
							disabled={solutionFilter.solutionsLoading}
						>
							{solutionFilter.solutionsLoading ? (
								<Option>Loading...</Option>
							) : unmanagedSolutions.length === 0 && managedSolutions.length === 0 ? (
								<Option>No solutions found</Option>
							) : (
								<>
									{unmanagedSolutions.length > 0 && (
										<OptionGroup label="Unmanaged">
											{unmanagedSolutions
												.filter((sol) =>
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase()),
												)
												.map((sol) => (
													<Option
														key={sol.solutionid}
														value={sol.solutionid}
														text={sol.friendlyname}
													>
														{sol.friendlyname}
													</Option>
												))}
										</OptionGroup>
									)}
									{managedSolutions.length > 0 && (
										<OptionGroup label="Managed">
											{managedSolutions
												.filter((sol) =>
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase()),
												)
												.map((sol) => (
													<Option
														key={sol.solutionid}
														value={sol.solutionid}
														text={sol.friendlyname}
													>
														{sol.friendlyname}
													</Option>
												))}
										</OptionGroup>
									)}
								</>
							)}
						</Combobox>
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
						placeholder={entityLoading ? "Loading entities..." : "Select an entity..."}
						value={entityQuery}
						onOptionSelect={onEntitySelect}
						onChange={(ev) => setEntityQuery(ev.target.value)}
						clearable
						disabled={entityLoading || availableEntities.length === 0}
					>
						{entityLoading ? (
							<Option>Loading...</Option>
						) : availableEntities.length === 0 ? (
							<Option>
								{effectiveScopeMode !== "all"
									? "Select solutions to see entities"
									: "No entities available"}
							</Option>
						) : (
							filteredEntityOptions
						)}
					</Combobox>
					{entityError && <div className={styles.errorText}>{entityError}</div>}
				</div>
				<Tooltip content="New query" relationship="label">
					<Button className={styles.newButton} appearance="primary" icon={<Add20Regular />} onClick={onNewQuery}>
						<span className={styles.newButtonLabel}>New</span>
					</Button>
				</Tooltip>
			</div>

			{/* Load View Picker - only show when entity is selected */}
			{selectedEntityMetadata && onViewLoad && (
				<div className={styles.viewPickerRow}>
					<LoadViewPicker
						selectedEntityMetadata={selectedEntityMetadata}
						onViewSelect={handleViewSelect}
						onViewClear={handleViewClear}
					/>
				</div>
			)}

		</div>
	);
}
