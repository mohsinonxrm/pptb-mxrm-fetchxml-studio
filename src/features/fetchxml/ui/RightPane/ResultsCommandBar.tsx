/**
 * Command bar for DataGrid with row actions
 * Similar to Power Apps Model Driven Apps view command bar
 */

import { useState, useCallback, type ReactElement } from "react";
import {
	Toolbar,
	ToolbarButton,
	ToolbarDivider,
	makeStyles,
	Tooltip,
	Menu,
	MenuTrigger,
	MenuPopover,
	MenuList,
	MenuItem,
	MenuItemRadio,
	MenuDivider,
	MenuGroup,
	MenuGroupHeader,
	SplitButton,
	MenuButton,
	type MenuButtonProps,
} from "@fluentui/react-components";
import {
	Open20Regular,
	Link20Regular,
	CheckmarkCircle20Regular,
	DismissCircle20Regular,
	Delete20Regular,
	ColumnTriple20Regular,
	Play20Regular,
	DocumentBulletList20Regular,
	Info20Regular,
	Stop20Regular,
	Copy20Regular,
} from "@fluentui/react-icons";
import { ExcelIcon } from "../../../../shared/components/ExcelIcon";
import { EditColumnsPanel } from "./EditColumnsPanel";
import {
	AddColumnsPanel,
	type AddColumnSelection,
	type RelatedEntityColumn,
} from "./AddColumnsPanel";
import type { LayoutColumn } from "../../model/layoutxml";
import type { AttributeMetadata, RelationshipMetadata, WorkflowInfo } from "../../api/pptbClient";
import type { CopyValueMode } from "./copyResults";

const useStyles = makeStyles({
	toolbar: {
		padding: "0",
	},

	splitButtonAppearance: {
		// Use subtle appearance for split buttons to match toolbar
	},
});

export interface CommandBarProps {
	selectedCount: number;
	onOpen?: () => void;
	onCopyUrl?: () => void;
	onActivate?: () => void;
	onDeactivate?: () => void;
	/** Why Activate is off for this entity (no statecode, close-message entity, no Write privilege, #143) */
	activateDisabledReason?: string;
	/** Why Deactivate is off for this entity (#143) */
	deactivateDisabledReason?: string;
	/** Delete selected records (or all if none selected) */
	onDelete?: () => void;
	/** Bulk delete (creates async job) */
	onBulkDelete?: () => void;
	/** Whether delete is available (user has privilege) */
	canDelete?: boolean;
	/** Whether bulk delete is available (user has privilege) */
	canBulkDelete?: boolean;
	/** Whether delete/bulk delete actions are disabled (e.g., aggregate query or 1-N relationship) */
	deleteDisabled?: boolean;
	/** Execute a specific workflow directly */
	onRunSpecificWorkflow?: (workflow: WorkflowInfo) => void;
	/** Whether workflow execution is available (user has privilege) */
	canRunWorkflow?: boolean;
	/** Whether workflow action is disabled (e.g., aggregate query or 1-N relationship) */
	workflowDisabled?: boolean;
	/**
	 * When set, the results have no record IDs (e.g. aggregate or distinct queries), so all record
	 * actions are disabled and this text explains why.
	 */
	recordActionsDisabledReason?: string;
	/**
	 * Why "Delete all" (no selection) is off: the query changed since these results loaded, or nothing
	 * ran yet, so it would delete records the grid doesn't show (#174). Selected records still work.
	 */
	deleteAllDisabledReason?: string;
	/** Fetch available workflows for the entity */
	onFetchWorkflows?: () => Promise<WorkflowInfo[]>;
	/** Export to Excel via Dataverse ExportToExcel API (requires saved view) */
	onExport?: () => void;
	/** Export to Excel locally using exceljs (no view required) */
	onExportLocal?: () => void;
	/** Whether server export is available (requires a saved view and export privilege) */
	canExport?: boolean;
	/** Whether export is in progress */
	isExporting?: boolean;
	/** Tooltip text to show when server export is disabled */
	exportDisabledReason?: string;
	/** Entity logical name */
	entityName?: string;
	/** Entity display name */
	entityDisplayName?: string;
	/** Current column layout */
	columns?: LayoutColumn[];
	/** Callback when columns are reordered */
	onReorderColumns?: (columns: LayoutColumn[]) => void;
	/** Callback when columns are removed */
	onRemoveColumns?: (columnKeys: string[]) => void;
	/** Available attributes from metadata for adding columns */
	availableAttributes?: AttributeMetadata[];
	/** Currently selected attributes in the query */
	selectedAttributes?: string[];
	/** Callback when user wants to add a column from root entity */
	onAddColumn?: (attributeName: string) => void;
	/** Callback when user wants to add columns from related entity */
	onAddRelatedColumns?: (columns: RelatedEntityColumn[]) => void;
	/** Lookup relationships (many-to-one) for related entity columns */
	lookupRelationships?: RelationshipMetadata[];
	/** One-to-many relationships for 1-N related entity columns */
	oneToManyRelationships?: RelationshipMetadata[];
	/** Whether relationship data is loading */
	isLoadingRelationships?: boolean;
	/** Callback to load attributes for a related entity */
	onLoadRelatedAttributes?: (entityLogicalName: string) => Promise<AttributeMetadata[]>;
	/** Why columns can't be added (e.g. aggregate queries), if they can't */
	addColumnsDisabledReason?: string;
	/** Related columns already returned: "entity|from|to|attribute" (Add columns leaves them out) */
	existingRelatedColumns?: ReadonlySet<string>;
	/** Callback to reset columns to default */
	onResetToDefault?: () => void;
	/** Retrieve All is fetching pages (#60) */
	isRetrieveAllInProgress?: boolean;
	/** Stop Retrieve All; the loaded rows stay (#60) */
	onStopRetrieveAll?: () => void;
	/** Rows loaded in the grid (the Copy menu's "All loaded rows") */
	loadedCount?: number;
	/** Copy the selected or all loaded rows to the clipboard (#50) */
	onCopyRows?: (scope: "selected" | "all", mode: CopyValueMode) => void;
	/** Quick search text (#23): the Copy menu then copies the matching rows */
	searchText?: string;
}

/** A toolbar button with a tooltip giving the reason it's off, when there is one */
function StateButton({
	label,
	ariaLabel,
	icon,
	onClick,
	disabled,
	reason,
}: {
	label: string;
	ariaLabel: string;
	icon: ReactElement;
	onClick?: () => void;
	disabled: boolean;
	reason?: string;
}) {
	const button = (
		<ToolbarButton appearance="subtle" icon={icon} disabled={disabled} disabledFocusable={!!reason} onClick={onClick} aria-label={ariaLabel}>
			{label}
		</ToolbarButton>
	);
	return reason ? (
		<Tooltip content={reason} relationship="description" withArrow>
			{button}
		</Tooltip>
	) : (
		button
	);
}

export function ResultsCommandBar({
	selectedCount,
	onOpen,
	onCopyUrl,
	onActivate,
	onDeactivate,
	activateDisabledReason,
	deactivateDisabledReason,
	onDelete,
	onBulkDelete,
	canDelete = false,
	canBulkDelete = false,
	deleteDisabled = false,
	onRunSpecificWorkflow,
	canRunWorkflow = false,
	workflowDisabled = false,
	recordActionsDisabledReason,
	deleteAllDisabledReason,
	onFetchWorkflows,
	onExport,
	onExportLocal,
	loadedCount = 0,
	onCopyRows,
	searchText = "",
	canExport = false,
	isExporting = false,
	exportDisabledReason,
	entityName,
	entityDisplayName,
	columns,
	onReorderColumns,
	onRemoveColumns,
	availableAttributes,
	selectedAttributes,
	onAddColumn,
	onAddRelatedColumns,
	addColumnsDisabledReason,
	existingRelatedColumns,
	lookupRelationships,
	oneToManyRelationships,
	isLoadingRelationships,
	onLoadRelatedAttributes,
	onResetToDefault,
	isRetrieveAllInProgress = false,
	onStopRetrieveAll,
}: CommandBarProps) {
	const styles = useStyles();
	const hasSelection = selectedCount > 0;
	const singleSelection = selectedCount === 1;
	const recordActionsUnavailable = !!recordActionsDisabledReason;
	const deleteAllBlocked = !hasSelection && !!deleteAllDisabledReason;
	const isDeleteDisabled = deleteDisabled || recordActionsUnavailable || deleteAllBlocked;
	const isWorkflowDisabled = workflowDisabled || recordActionsUnavailable;
	// Activate / Deactivate (#143): need a selection, record IDs and the record-action rules delete has
	const stateButtonsDisabled = !hasSelection || isDeleteDisabled;
	const [editPanelOpen, setEditPanelOpen] = useState(false);
	const [addPanelOpen, setAddPanelOpen] = useState(false);
	// Copy menu: values as shown in the grid, or raw (#50)
	const [copyMode, setCopyMode] = useState<CopyValueMode>("shown");

	// Workflow menu state
	const [workflows, setWorkflows] = useState<WorkflowInfo[]>([]);
	const [isLoadingWorkflows, setIsLoadingWorkflows] = useState(false);
	const [workflowsLoaded, setWorkflowsLoaded] = useState(false);

	// Load workflows when menu opens (lazy loading)
	const handleWorkflowMenuOpen = useCallback(async () => {
		if (workflowsLoaded || isLoadingWorkflows || !onFetchWorkflows) return;

		setIsLoadingWorkflows(true);
		try {
			const wfs = await onFetchWorkflows();
			setWorkflows(wfs);
			setWorkflowsLoaded(true);
		} catch (err) {
			console.error("Failed to load workflows:", err);
		} finally {
			setIsLoadingWorkflows(false);
		}
	}, [workflowsLoaded, isLoadingWorkflows, onFetchWorkflows]);

	const handleEditColumnsClick = useCallback(() => {
		setEditPanelOpen(true);
	}, []);

	const handleEditPanelClose = useCallback(() => {
		setEditPanelOpen(false);
	}, []);

	const handleAddPanelClose = useCallback(() => {
		setAddPanelOpen(false);
	}, []);

	const handleEditApply = useCallback(
		(updatedColumns: LayoutColumn[]) => {
			// Remove the attributes behind removed columns, all at once
			if (columns && onRemoveColumns) {
				const updatedNames = new Set(updatedColumns.map((c) => c.name));
				const removed = columns.filter((col) => !updatedNames.has(col.name)).map((col) => col.name);
				if (removed.length > 0) onRemoveColumns(removed);
			}
			// Apply reordering
			onReorderColumns?.(updatedColumns);
			setEditPanelOpen(false);
		},
		[columns, onReorderColumns, onRemoveColumns]
	);

	const handleAddApply = useCallback(
		(selection: AddColumnSelection) => {
			// Add root entity attributes
			for (const attrName of selection.rootAttributes) {
				onAddColumn?.(attrName);
			}
			// Add related entity columns
			if (selection.relatedColumns.length > 0 && onAddRelatedColumns) {
				onAddRelatedColumns(selection.relatedColumns);
			}
			setAddPanelOpen(false);
		},
		[onAddColumn, onAddRelatedColumns]
	);

	const handleOpenAddFromEdit = useCallback(() => {
		setAddPanelOpen(true);
	}, []);

	return (
		<>
			<Toolbar className={styles.toolbar} size="medium" aria-label="Record actions">
				{isRetrieveAllInProgress && onStopRetrieveAll && (
					<>
						<ToolbarButton appearance="subtle" icon={<Stop20Regular />} onClick={onStopRetrieveAll}>
							Stop
						</ToolbarButton>
						<ToolbarDivider />
					</>
				)}
				<ToolbarButton
					appearance="subtle"
					icon={<Open20Regular />}
					disabled={!singleSelection || recordActionsUnavailable}
					onClick={onOpen}
					aria-label="Open record"
				>
					Open
				</ToolbarButton>

				<ToolbarButton
					appearance="subtle"
					icon={<Link20Regular />}
					disabled={!singleSelection || recordActionsUnavailable}
					onClick={onCopyUrl}
					aria-label="Copy record URL"
				>
					Copy URL
				</ToolbarButton>

				{recordActionsDisabledReason && (
					<Tooltip content={recordActionsDisabledReason} relationship="description" withArrow>
						<ToolbarButton
							appearance="transparent"
							icon={<Info20Regular />}
						>
							Record actions unavailable
						</ToolbarButton>
					</Tooltip>
				)}

				<ToolbarDivider />

				{/* Activate / Deactivate (#143); the tooltip says why one is off for this entity */}
				<StateButton
					label="Activate"
					ariaLabel="Activate selected records"
					icon={<CheckmarkCircle20Regular />}
					onClick={onActivate}
					disabled={stateButtonsDisabled || !onActivate || !!activateDisabledReason}
					reason={activateDisabledReason}
				/>

				<StateButton
					label="Deactivate"
					ariaLabel="Deactivate selected records"
					icon={<DismissCircle20Regular />}
					onClick={onDeactivate}
					disabled={stateButtonsDisabled || !onDeactivate || !!deactivateDisabledReason}
					reason={deactivateDisabledReason}
				/>

				<ToolbarDivider />

				{/* Same pattern as "Record actions unavailable": the reason is in the tooltip (#174) */}
				{canDelete && deleteAllBlocked && !recordActionsDisabledReason && (
					<Tooltip content={deleteAllDisabledReason} relationship="description" withArrow>
						<ToolbarButton appearance="transparent" icon={<Info20Regular />}>
							Delete all unavailable
						</ToolbarButton>
					</Tooltip>
				)}
				{canDelete && (
					<Menu positioning="below-end">
						<MenuTrigger disableButtonEnhancement>
							{(triggerProps: MenuButtonProps) => (
								<SplitButton
									appearance="subtle"
									menuButton={triggerProps}
									primaryActionButton={{
										onClick: onDelete,
										disabled: isDeleteDisabled,
									}}
									disabled={isDeleteDisabled}
									icon={<Delete20Regular />}
									aria-label={
										hasSelection ? "Delete selected records" : "Delete all records from view"
									}
								>
									{hasSelection ? "Delete" : "Delete all"}
								</SplitButton>
							)}
						</MenuTrigger>
						<MenuPopover>
							<MenuList>
								<MenuItem icon={<Delete20Regular />} onClick={onDelete} disabled={isDeleteDisabled}>
									Delete{hasSelection ? ` (${selectedCount})` : " All"}
								</MenuItem>
								{canBulkDelete && (
									<MenuItem
										icon={<DocumentBulletList20Regular />}
										onClick={onBulkDelete}
										disabled={isDeleteDisabled}
									>
										Bulk delete (job)
									</MenuItem>
								)}
							</MenuList>
						</MenuPopover>
					</Menu>
				)}

				{canRunWorkflow && (
					<Menu
						positioning="below-end"
						onOpenChange={(_, data) => data.open && handleWorkflowMenuOpen()}
					>
						<MenuTrigger disableButtonEnhancement>
							<MenuButton
								appearance="subtle"
								icon={<Play20Regular />}
								disabled={!hasSelection || isWorkflowDisabled}
								aria-label="Run workflow on selected records"
							>
								Run workflow
							</MenuButton>
						</MenuTrigger>
						<MenuPopover>
							<MenuList>
								{isLoadingWorkflows && <MenuItem disabled>Loading workflows...</MenuItem>}
								{!isLoadingWorkflows && workflows.length === 0 && workflowsLoaded && (
									<MenuItem disabled>No workflows available</MenuItem>
								)}
								{!isLoadingWorkflows &&
									workflows.map((wf) => (
										<MenuItem
											key={wf.workflowid}
											icon={<Play20Regular />}
											onClick={() => onRunSpecificWorkflow?.(wf)}
										>
											{wf.name}
										</MenuItem>
									))}
							</MenuList>
						</MenuPopover>
					</Menu>
				)}

				<ToolbarDivider />

				{/* Copy rows as tab-separated values with headers; pastes into Excel (#50) */}
				<Menu
					checkedValues={{ copyValues: [copyMode] }}
					onCheckedValueChange={(_, data) => setCopyMode(data.checkedItems[0] as CopyValueMode)}
				>
					<MenuTrigger disableButtonEnhancement>
						<MenuButton appearance="subtle" icon={<Copy20Regular />} disabled={loadedCount === 0}>
							Copy
						</MenuButton>
					</MenuTrigger>
					<MenuPopover>
						<MenuList>
							<MenuItem disabled={selectedCount === 0} onClick={() => onCopyRows?.("selected", copyMode)}>
								Selected rows ({selectedCount})
							</MenuItem>
							<MenuItem onClick={() => onCopyRows?.("all", copyMode)}>
								{searchText.trim() ? "All matching rows" : "All loaded rows"} ({loadedCount})
							</MenuItem>
							<MenuDivider />
							<MenuGroup>
								<MenuGroupHeader>Values</MenuGroupHeader>
								<MenuItemRadio name="copyValues" value="shown" persistOnClick>
									As shown
								</MenuItemRadio>
								<MenuItemRadio name="copyValues" value="raw" persistOnClick>
									Raw (GUIDs, numbers, ISO dates)
								</MenuItemRadio>
							</MenuGroup>
						</MenuList>
					</MenuPopover>
				</Menu>

				{/* Export Menu Button */}
				<Menu>
					<MenuTrigger disableButtonEnhancement>
						<Tooltip
							content={isExporting ? "Export in progress..." : "Export to Excel"}
							relationship="description"
						>
							<MenuButton
								appearance="subtle"
								icon={<ExcelIcon />}
								disabled={isExporting}
								aria-label="Export to Excel"
							>
								{isExporting ? "Exporting..." : "Export to Excel"}
							</MenuButton>
						</Tooltip>
					</MenuTrigger>
					<MenuPopover>
						<MenuList>
							<Tooltip
								content={
									canExport
										? "Export using Dataverse API (all records)"
										: exportDisabledReason || "Save as a view first to enable server export"
								}
								relationship="description"
								positioning="before"
							>
								<MenuItem onClick={onExport} disabled={!canExport || isExporting}>
									Export (Server)
								</MenuItem>
							</Tooltip>
							<MenuItem onClick={onExportLocal} disabled={isExporting}>
								Export (Local)
							</MenuItem>
						</MenuList>
					</MenuPopover>
				</Menu>

				<ToolbarDivider />

				<ToolbarButton
					appearance="subtle"
					icon={<ColumnTriple20Regular />}
					onClick={handleEditColumnsClick}
					aria-label="Edit columns"
				>
					Edit columns
				</ToolbarButton>
			</Toolbar>

			{/* Edit Columns Panel */}
			<EditColumnsPanel
				open={editPanelOpen}
				columns={columns || []}
				entityDisplayName={entityDisplayName}
				onClose={handleEditPanelClose}
				onApply={handleEditApply}
				onAddColumns={handleOpenAddFromEdit}
				onResetToDefault={onResetToDefault}
				addColumnsDisabledReason={addColumnsDisabledReason}
			/>

			{/* Add Columns Panel */}
			<AddColumnsPanel
				open={addPanelOpen}
				entityDisplayName={entityDisplayName}
				entityLogicalName={entityName}
				availableAttributes={availableAttributes}
				selectedAttributes={selectedAttributes}
				lookupRelationships={lookupRelationships}
				oneToManyRelationships={oneToManyRelationships}
				isLoadingRelationships={isLoadingRelationships}
				onLoadRelatedAttributes={onLoadRelatedAttributes}
				existingRelatedColumns={existingRelatedColumns}
				onClose={handleAddPanelClose}
				onApply={handleAddApply}
			/>
		</>
	);
}
