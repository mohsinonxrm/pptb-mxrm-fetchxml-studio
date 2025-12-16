/**
 * Command bar for DataGrid with row actions
 * Similar to Power Apps Model Driven Apps view command bar
 */

import { useState, useCallback } from "react";
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
	onRemoveColumn?: (columnName: string) => void;
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
	/** Callback to reset columns to default */
	onResetToDefault?: () => void;
}

export function ResultsCommandBar({
	selectedCount,
	onOpen,
	onCopyUrl,
	onActivate,
	onDeactivate,
	onDelete,
	onBulkDelete,
	canDelete = false,
	canBulkDelete = false,
	deleteDisabled = false,
	onRunSpecificWorkflow,
	canRunWorkflow = false,
	workflowDisabled = false,
	onFetchWorkflows,
	onExport,
	onExportLocal,
	canExport = false,
	isExporting = false,
	exportDisabledReason,
	entityName,
	entityDisplayName,
	columns,
	onReorderColumns,
	onRemoveColumn,
	availableAttributes,
	selectedAttributes,
	onAddColumn,
	onAddRelatedColumns,
	lookupRelationships,
	oneToManyRelationships,
	isLoadingRelationships,
	onLoadRelatedAttributes,
	onResetToDefault,
}: CommandBarProps) {
	const styles = useStyles();
	const hasSelection = selectedCount > 0;
	const singleSelection = selectedCount === 1;
	const [editPanelOpen, setEditPanelOpen] = useState(false);
	const [addPanelOpen, setAddPanelOpen] = useState(false);

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
			// Detect removed columns
			if (columns && onRemoveColumn) {
				const updatedNames = new Set(updatedColumns.map((c) => c.name));
				for (const col of columns) {
					if (!updatedNames.has(col.name)) {
						onRemoveColumn(col.name);
					}
				}
			}
			// Apply reordering
			onReorderColumns?.(updatedColumns);
			setEditPanelOpen(false);
		},
		[columns, onReorderColumns, onRemoveColumn]
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
				<ToolbarButton
					appearance="subtle"
					icon={<Open20Regular />}
					disabled={!singleSelection}
					onClick={onOpen}
					aria-label="Open record"
				>
					Open
				</ToolbarButton>

				<ToolbarButton
					appearance="subtle"
					icon={<Link20Regular />}
					disabled={!singleSelection}
					onClick={onCopyUrl}
					aria-label="Copy record URL"
				>
					Copy URL
				</ToolbarButton>

				<ToolbarDivider />

				{/* Activate/Deactivate buttons - always shown, disabled for now */}
				<ToolbarButton
					appearance="subtle"
					icon={<CheckmarkCircle20Regular />}
					disabled={true}
					onClick={onActivate}
					aria-label="Activate selected records"
				>
					Activate
				</ToolbarButton>

				<ToolbarButton
					appearance="subtle"
					icon={<DismissCircle20Regular />}
					disabled={true}
					onClick={onDeactivate}
					aria-label="Deactivate selected records"
				>
					Deactivate
				</ToolbarButton>

				<ToolbarDivider />

				{canDelete && (
					<Menu positioning="below-end">
						<MenuTrigger disableButtonEnhancement>
							{(triggerProps: MenuButtonProps) => (
								<SplitButton
									appearance="subtle"
									menuButton={triggerProps}
									primaryActionButton={{
										onClick: onDelete,
										disabled: deleteDisabled,
									}}
									disabled={deleteDisabled}
									icon={<Delete20Regular />}
									aria-label={
										hasSelection ? "Delete selected records" : "Delete all records from view"
									}
								>
									{hasSelection ? "Delete" : "Delete All"}
								</SplitButton>
							)}
						</MenuTrigger>
						<MenuPopover>
							<MenuList>
								<MenuItem icon={<Delete20Regular />} onClick={onDelete} disabled={deleteDisabled}>
									Delete{hasSelection ? ` (${selectedCount})` : " All"}
								</MenuItem>
								{canBulkDelete && (
									<MenuItem
										icon={<DocumentBulletList20Regular />}
										onClick={onBulkDelete}
										disabled={deleteDisabled}
									>
										Bulk Delete (Job)
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
								disabled={!hasSelection || workflowDisabled}
								aria-label="Run workflow on selected records"
							>
								Run Workflow
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
				onClose={handleAddPanelClose}
				onApply={handleAddApply}
			/>
		</>
	);
}
