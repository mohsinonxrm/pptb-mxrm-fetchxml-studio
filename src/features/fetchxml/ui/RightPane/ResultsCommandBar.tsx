/**
 * Command bar for DataGrid with row actions
 * Similar to Power Apps Model Driven Apps view command bar
 */

import { useMemo, useState, useCallback } from "react";
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
	ArrowExport20Regular,
	ColumnTriple20Regular,
	Play20Regular,
	DocumentBulletList20Regular,
} from "@fluentui/react-icons";
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
	/** Whether delete is available */
	canDelete?: boolean;
	/** Whether bulk delete is available */
	canBulkDelete?: boolean;
	/** Execute a specific workflow directly */
	onRunSpecificWorkflow?: (workflow: WorkflowInfo) => void;
	/** Whether workflow execution is available */
	canRunWorkflow?: boolean;
	/** Fetch available workflows for the entity */
	onFetchWorkflows?: () => Promise<WorkflowInfo[]>;
	onExport?: () => void;
	/** Whether export is available (requires a saved view and export privilege) */
	canExport?: boolean;
	/** Whether export is in progress */
	isExporting?: boolean;
	/** Tooltip text to show when export is disabled */
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
	onRunSpecificWorkflow,
	canRunWorkflow = false,
	onFetchWorkflows,
	onExport,
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

	// Determine if entity supports statecode/statuscode (most common pattern)
	const supportsActivation = useMemo(() => {
		if (!entityName) return false;
		// Most entities support activation, but some don't
		const noActivationEntities = [
			"activitypointer",
			"annotation",
			"note",
			"connection",
			"connectionrole",
			"savedquery",
			"userquery",
		];
		return !noActivationEntities.includes(entityName.toLowerCase());
	}, [entityName]);

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

				{supportsActivation && (
					<>
						<ToolbarButton
							appearance="subtle"
							icon={<CheckmarkCircle20Regular />}
							disabled={!hasSelection}
							onClick={onActivate}
							aria-label="Activate selected records"
						>
							Activate
						</ToolbarButton>

						<ToolbarButton
							appearance="subtle"
							icon={<DismissCircle20Regular />}
							disabled={!hasSelection}
							onClick={onDeactivate}
							aria-label="Deactivate selected records"
						>
							Deactivate
						</ToolbarButton>

						<ToolbarDivider />
					</>
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
									}}
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
								<MenuItem icon={<Delete20Regular />} onClick={onDelete}>
									Delete{hasSelection ? ` (${selectedCount})` : " All"}
								</MenuItem>
								{canBulkDelete && (
									<MenuItem icon={<DocumentBulletList20Regular />} onClick={onBulkDelete}>
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
								disabled={!hasSelection}
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

				<Tooltip
					content={
						canExport
							? "Export to Excel"
							: exportDisabledReason || "Save as a view first to enable export"
					}
					relationship="description"
				>
					<ToolbarButton
						appearance="subtle"
						icon={<ArrowExport20Regular />}
						onClick={onExport}
						disabled={!canExport || isExporting}
						aria-label="Export to Excel"
					>
						{isExporting ? "Exporting..." : "Export"}
					</ToolbarButton>
				</Tooltip>

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
				isLoadingRelationships={isLoadingRelationships}
				onLoadRelatedAttributes={onLoadRelatedAttributes}
				onClose={handleAddPanelClose}
				onApply={handleAddApply}
			/>
		</>
	);
}
