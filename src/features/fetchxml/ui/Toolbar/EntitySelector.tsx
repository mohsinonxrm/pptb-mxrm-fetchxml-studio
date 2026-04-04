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
	Dialog,
	DialogSurface,
	DialogBody,
	DialogTitle,
	DialogContent,
	DialogActions,
	type ComboboxProps,
} from "@fluentui/react-components";
import { Add20Regular, LockClosed20Regular } from "@fluentui/react-icons";
import { useAccessMode } from "../../../../shared/hooks/useAccessMode";
import type { EntityScopeMode } from "../../model/displaySettings";
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

		// Apply Advanced Find filter locally — no re-fetch needed when toggled
		if (advancedFindOnly) {
			entities = entities.filter((e) => e.IsValidForAdvancedFind !== false);
		}

		return entities;
	}, [
		effectiveScopeMode,
		advancedFindOnly,
		publisherFilter.entities,
		solutionFilter.entities,
		allEntities,
	]);

	// Get full entity metadata for the selected entity (needed for LoadViewPicker)
	const selectedEntityMetadata = useMemo(() => {
		if (!selectedEntity) return null;
		return availableEntities.find((e) => e.LogicalName === selectedEntity) || null;
	}, [selectedEntity, availableEntities]);

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

	const onPublisherSelect: ComboboxProps["onOptionSelect"] = async (_e, data) => {
		const newPublisherIds = data.selectedOptions;

		// Check if this change would invalidate the current entity
		if (selectedEntity && publisherFilter.selectedPublisherIds.length > 0) {
			// If deselecting publishers, check if entity would still be available
			const isDeselecting = newPublisherIds.length < publisherFilter.selectedPublisherIds.length;

			if (isDeselecting) {
				// Find which publishers are being removed
				const removedPublisherIds = publisherFilter.selectedPublisherIds.filter(
					(id) => !newPublisherIds.includes(id),
				);

				// Find solutions under removed publishers
				const removedPublisherSolutionIds = publisherFilter.solutions
					.filter(
						(sol) => sol._publisherid_value && removedPublisherIds.includes(sol._publisherid_value),
					)
					.map((sol) => sol.solutionid);

				// Check if any of the currently selected solutions are from removed publishers
				const wouldLoseSolutions = publisherFilter.selectedSolutionIds.some((solId) =>
					removedPublisherSolutionIds.includes(solId),
				);

				// If we'd lose solutions that might contain the entity, show confirmation
				if (wouldLoseSolutions) {
					setPendingPublisherIds(newPublisherIds);
					setConfirmDialogType("publisher");
					setShowConfirmDialog(true);
					setPublisherQuery("");
					return;
				}
			}
		}

		// No validation needed, proceed with update
		publisherFilter.updateSelectedPublishers(newPublisherIds);
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

	const onSolutionSelect: ComboboxProps["onOptionSelect"] = async (_e, data) => {
		const newSolutionIds = data.selectedOptions;

		// Check if this change would invalidate the current entity
		if (selectedEntity && currentSelectedSolutionIds.length > 0) {
			// If deselecting solutions, check if entity would still be available
			const isDeselecting = newSolutionIds.length < currentSelectedSolutionIds.length;

			if (isDeselecting) {
				// Find which solutions are being removed
				const removedSolutionIds = currentSelectedSolutionIds.filter(
					(id) => !newSolutionIds.includes(id),
				);

				// Check if current entity exists in the remaining solutions
				const remainingEntityExists = availableEntities.some(
					(entity) => entity.LogicalName === selectedEntity,
				);

				// Only show confirmation if entity would be removed
				if (remainingEntityExists && removedSolutionIds.length > 0) {
					// We need to check if entity exists in remaining solutions after removal
					// For now, show confirmation dialog and let validation happen after
					setPendingSolutionIds(newSolutionIds);
					setConfirmDialogType("solution");
					setShowConfirmDialog(true);
					setSolutionQuery("");
					return;
				}
			}
		}

		// No validation needed, proceed with update
		if (effectiveScopeMode === "publisher-solution") {
			publisherFilter.updateSelectedSolutions(newSolutionIds);
		} else if (effectiveScopeMode === "solution-only") {
			solutionFilter.updateSelectedSolutions(newSolutionIds);
		}
		setSolutionQuery("");
	};

	const handleConfirmPublisherChange = () => {
		// User confirmed - proceed with publisher change, cascades to solutions and entity
		publisherFilter.updateSelectedPublishers(pendingPublisherIds);
		setShowConfirmDialog(false);
		setPendingPublisherIds([]);
	};

	const handleConfirmSolutionChange = () => {
		// User confirmed - proceed with solution change, entity will auto-clear via useEffect
		if (effectiveScopeMode === "publisher-solution") {
			publisherFilter.updateSelectedSolutions(pendingSolutionIds);
		} else if (effectiveScopeMode === "solution-only") {
			solutionFilter.updateSelectedSolutions(pendingSolutionIds);
		}
		setShowConfirmDialog(false);
		setPendingSolutionIds([]);
	};

	const handleCancelChange = () => {
		// User cancelled - just close dialog, keep current selection
		setShowConfirmDialog(false);
		setPendingPublisherIds([]);
		setPendingSolutionIds([]);
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

	// Validate entity when available entities change due to the user actively changing
	// their solution filter. Only clears when an active solution filter is in place —
	// prevents resetting entities typed manually in the XML editor.
	useEffect(() => {
		if (!selectedEntity) return;

		// Only auto-clear if the user has explicitly filtered to specific solutions.
		// Without an active filter the entity may have been set via the XML editor;
		// clearing it would be surprising and unwanted.
		const hasActiveSolutionFilter =
			(effectiveScopeMode === "publisher-solution" &&
				publisherFilter.selectedSolutionIds.length > 0) ||
			(effectiveScopeMode === "solution-only" && solutionFilter.selectedSolutionIds.length > 0);

		if (!hasActiveSolutionFilter) return;

		const isEntityStillAvailable = availableEntities.some(
			(entity) => entity.LogicalName === selectedEntity,
		);

		if (!isEntityStillAvailable) {
			onEntityChange("");
			onNewQuery();
		}
	}, [
		availableEntities,
		selectedEntity,
		onEntityChange,
		onNewQuery,
		effectiveScopeMode,
		publisherFilter.selectedSolutionIds,
		solutionFilter.selectedSolutionIds,
	]);

	// Entity search query with filteringhanges that would invalidate entity
	const [showConfirmDialog, setShowConfirmDialog] = useState(false);
	const [confirmDialogType, setConfirmDialogType] = useState<"publisher" | "solution">("solution");
	const [pendingPublisherIds, setPendingPublisherIds] = useState<string[]>([]);
	const [pendingSolutionIds, setPendingSolutionIds] = useState<string[]>([]);

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
				<Button
					appearance="primary"
					icon={<Add20Regular />}
					onClick={onNewQuery}
					title="Create a new query"
				>
					New
				</Button>
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

			{/* Confirmation Dialog for Publisher/Solution Change */}
			<Dialog
				open={showConfirmDialog}
				onOpenChange={(_, data) => !data.open && handleCancelChange()}
			>
				<DialogSurface>
					<DialogBody>
						<DialogTitle>
							{confirmDialogType === "publisher"
								? "Confirm Publisher Change"
								: "Confirm Solution Change"}
						</DialogTitle>
						<DialogContent>
							<p>
								{confirmDialogType === "publisher"
									? "Removing publishers will also remove their associated solutions, which may reset the FetchXML tree and entity selection as"
									: "Removing solutions will reset the FetchXML tree and entity selection as"}
								<strong> {selectedEntity}</strong> may no longer be available
								{confirmDialogType === "publisher"
									? " in the remaining publishers."
									: " in the remaining selected solutions."}
							</p>
							<p>This action cannot be undone. Do you want to proceed?</p>
						</DialogContent>
						<DialogActions>
							<Button
								appearance="primary"
								onClick={
									confirmDialogType === "publisher"
										? handleConfirmPublisherChange
										: handleConfirmSolutionChange
								}
							>
								{confirmDialogType === "publisher"
									? "Yes, Update Publishers"
									: "Yes, Update Solutions"}
							</Button>
							<Button appearance="secondary" onClick={handleCancelChange}>
								Cancel
							</Button>
						</DialogActions>
					</DialogBody>
				</DialogSurface>
			</Dialog>
		</div>
	);
}
