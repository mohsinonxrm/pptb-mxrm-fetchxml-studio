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
	},
	filtersRow: {
		display: "flex",
		gap: "12px",
		alignItems: "flex-end",
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
}

export function EntitySelector({
	selectedEntity,
	onEntityChange,
	onNewQuery,
	onViewLoad,
}: EntitySelectorProps) {
	const styles = useStyles();
	const publisherComboId = useId("publisher-combobox");
	const solutionComboId = useId("solution-combobox");
	const entityComboId = useId("entity-combobox");

	// Access mode detection
	const {
		loading: accessLoading,
		fullFilterMode,
		solutionsOnlyMode,
		publishersOnlyMode,
		metadataOnlyMode,
		noAccessMode,
	} = useAccessMode();

	// Full filter mode (Publisher → Solution → Entity)
	const publisherFilter = usePublisherFilter();

	// Solutions-only mode (Solution → Entity)
	const solutionFilter = useSolutionFilter();

	// Metadata-only mode (all AF-valid entities) or Publishers-only mode
	const { loadEntities } = useLazyMetadata();
	const [allEntities, setAllEntities] = useState<EntityMetadata[]>([]);
	const [allEntitiesLoading, setAllEntitiesLoading] = useState(false);

	// Load all entities for metadata-only or publishers-only mode
	useEffect(() => {
		if (!metadataOnlyMode && !publishersOnlyMode) return;

		setAllEntitiesLoading(true);
		loadEntities(true)
			.then((entities) => setAllEntities(entities))
			.catch((err) => console.error("Failed to load entities:", err))
			.finally(() => setAllEntitiesLoading(false));
	}, [metadataOnlyMode, publishersOnlyMode, loadEntities]);

	// Determine available entities based on mode
	const availableEntities = useMemo(() => {
		let entities: EntityMetadata[] = [];
		if (fullFilterMode) {
			entities = publisherFilter.entities;
		} else if (solutionsOnlyMode) {
			entities = solutionFilter.entities;
		} else if (metadataOnlyMode || publishersOnlyMode) {
			entities = allEntities;
		}

		console.log("[EntitySelector] Available Entities Updated:", {
			mode: fullFilterMode ? "Full Filter" : solutionsOnlyMode ? "Solutions Only" : "Other",
			selectedSolutionIds: fullFilterMode
				? publisherFilter.selectedSolutionIds
				: solutionFilter.selectedSolutionIds,
			entityCount: entities.length,
			entityNames: entities.map((e) => e.LogicalName).slice(0, 10),
		});

		return entities;
	}, [
		fullFilterMode,
		solutionsOnlyMode,
		metadataOnlyMode,
		publishersOnlyMode,
		publisherFilter.entities,
		solutionFilter.entities,
		allEntities,
		publisherFilter.selectedSolutionIds,
		solutionFilter.selectedSolutionIds,
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
		[onViewLoad]
	);

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
			(opt) => typeof opt.children === "string" && opt.children.toLowerCase().includes(lowerQuery)
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
					(id) => !newPublisherIds.includes(id)
				);

				// Find solutions under removed publishers
				const removedPublisherSolutionIds = publisherFilter.solutions
					.filter(
						(sol) => sol._publisherid_value && removedPublisherIds.includes(sol._publisherid_value)
					)
					.map((sol) => sol.solutionid);

				// Check if any of the currently selected solutions are from removed publishers
				const wouldLoseSolutions = publisherFilter.selectedSolutionIds.some((solId) =>
					removedPublisherSolutionIds.includes(solId)
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
	const currentSolutions = fullFilterMode ? publisherFilter.solutions : solutionFilter.solutions;

	const currentSelectedSolutionIds = fullFilterMode
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
					(id) => !newSolutionIds.includes(id)
				);

				// Check if current entity exists in the remaining solutions
				const remainingEntityExists = availableEntities.some(
					(entity) => entity.LogicalName === selectedEntity
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
		if (fullFilterMode) {
			publisherFilter.updateSelectedSolutions(newSolutionIds);
		} else if (solutionsOnlyMode) {
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
		if (fullFilterMode) {
			publisherFilter.updateSelectedSolutions(pendingSolutionIds);
		} else if (solutionsOnlyMode) {
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

	// Validate entity when available entities change (solution selection changes)
	useEffect(() => {
		// Only validate if we have a selected entity
		if (!selectedEntity) return;

		// If available entities is empty, clear entity and tree
		if (availableEntities.length === 0) {
			onEntityChange(""); // Clear entity selection
			onNewQuery(); // Reset fetch tree
			return;
		}

		// Check if selected entity is still available
		const isEntityStillAvailable = availableEntities.some(
			(entity) => entity.LogicalName === selectedEntity
		);

		// If entity is no longer available, reset it and the fetch tree
		if (!isEntityStillAvailable) {
			onEntityChange(""); // Clear entity selection
			onNewQuery(); // Reset fetch tree
		}
	}, [availableEntities, selectedEntity, onEntityChange, onNewQuery]);

	// Entity search query with filteringhanges that would invalidate entity
	const [showConfirmDialog, setShowConfirmDialog] = useState(false);
	const [confirmDialogType, setConfirmDialogType] = useState<"publisher" | "solution">("solution");
	const [pendingPublisherIds, setPendingPublisherIds] = useState<string[]>([]);
	const [pendingSolutionIds, setPendingSolutionIds] = useState<string[]>([]);

	// Entity search query with filtering
	const [entityQuery, setEntityQuery] = useState<string>("");

	const entityOptions = useMemo(() => {
		return availableEntities.map((entity) => ({
			value: entity.LogicalName,
			children: entity.DisplayName?.UserLocalizedLabel?.Label || entity.LogicalName,
		}));
	}, [availableEntities]);

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
	const entityLoading = fullFilterMode
		? publisherFilter.entitiesLoading
		: solutionsOnlyMode
		? solutionFilter.entitiesLoading
		: allEntitiesLoading;

	const entityError = fullFilterMode
		? publisherFilter.entitiesError
		: solutionsOnlyMode
		? solutionFilter.entitiesError
		: null;

	return (
		<div className={styles.container}>
			{/* Full Filter Mode: Publisher + Solution filters */}
			{fullFilterMode && (
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
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase())
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
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase())
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

			{/* Solutions-Only Mode: Solution filter with disabled Publisher */}
			{solutionsOnlyMode && (
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
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase())
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
													sol.friendlyname.toLowerCase().includes(solutionQuery.toLowerCase())
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

			{/* Publishers-only mode: Disabled Publisher with tooltip + direct entity loading */}
			{publishersOnlyMode && (
				<div className={styles.filtersRow}>
					{/* Publisher - disabled with tooltip */}
					<div className={styles.field}>
						<Tooltip
							content="Solution filtering requires prvReadSolution privilege"
							relationship="description"
						>
							<label id={publisherComboId} className={`${styles.label} ${styles.disabledLabel}`}>
								Publishers <LockClosed20Regular fontSize={14} />
							</label>
						</Tooltip>
						<Combobox
							aria-labelledby={publisherComboId}
							placeholder="No access to solutions"
							disabled
						>
							<Option>Requires prvReadSolution privilege</Option>
						</Combobox>
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
								{fullFilterMode || solutionsOnlyMode
									? "Select solutions to see entities"
									: publishersOnlyMode
									? "Loading all entities..."
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
