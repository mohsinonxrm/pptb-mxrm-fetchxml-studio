/**
 * The results command bar (#232): record actions and the results' own commands on the shared
 * OverflowToolbar, in the order of a model-driven view's command bar. Commands that don't fit move
 * into "More commands" by priority (D: Open 50, Delete 45, Copy URL 40, Copy and Edit columns 30,
 * Deactivate / Close 25, Run workflow 22, Activate 20, Export 15, Density 5); Stop, while Retrieve
 * all runs, never does.
 *
 * A command that can't be used right now stays focusable with its reason (resultsBarRules). One the
 * user lacks the privilege for is hidden, once the check has finished (owner decision 1).
 */

import { Fragment, useState, useCallback } from "react";
import {
	makeStyles,
	Tooltip,
	MenuList,
	MenuItem,
	MenuItemRadio,
	MenuDivider,
	MenuGroup,
	MenuGroupHeader,
	tokens,
} from "@fluentui/react-components";
import {
	Open20Filled,
	Open20Regular,
	Link20Filled,
	Link20Regular,
	CheckmarkCircle20Filled,
	CheckmarkCircle20Regular,
	DismissCircle20Filled,
	DismissCircle20Regular,
	CircleOff20Filled,
	CircleOff20Regular,
	Delete20Filled,
	Delete20Regular,
	ColumnTriple20Filled,
	ColumnTriple20Regular,
	Play20Filled,
	Play20Regular,
	DocumentBulletList20Regular,
	Stop20Filled,
	Stop20Regular,
	Copy20Filled,
	Copy20Regular,
	TextDensity20Filled,
	TextDensity20Regular,
	bundleIcon,
} from "@fluentui/react-icons";
import { ExcelIcon } from "../../../../shared/components/ExcelIcon";
import { OverflowToolbar, type BarCommand } from "../../../../shared/components/commandBar";
import { EditColumnsPanel } from "./EditColumnsPanel";
import {
	AddColumnsPanel,
	type AddColumnSelection,
	type RelatedEntityColumn,
} from "./AddColumnsPanel";
import type { LayoutColumn } from "../../model/layoutxml";
import type { AttributeMetadata, RelationshipMetadata, WorkflowInfo } from "../../api/pptbClient";
import { useEntityWorkflows } from "./useEntityWorkflows";
import type { CopyValueMode } from "./copyResults";
import { GRID_DENSITY_LABELS, type GridDensity } from "../../model/displaySettings";
import { resultsBarReasons } from "./resultsBarRules";

const OpenIcon = bundleIcon(Open20Filled, Open20Regular);
const LinkIcon = bundleIcon(Link20Filled, Link20Regular);
const ActivateIcon = bundleIcon(CheckmarkCircle20Filled, CheckmarkCircle20Regular);
const DeactivateIcon = bundleIcon(DismissCircle20Filled, DismissCircle20Regular);
const CloseIcon = bundleIcon(CircleOff20Filled, CircleOff20Regular);
const DeleteIcon = bundleIcon(Delete20Filled, Delete20Regular);
const ColumnsIcon = bundleIcon(ColumnTriple20Filled, ColumnTriple20Regular);
const WorkflowIcon = bundleIcon(Play20Filled, Play20Regular);
const StopIcon = bundleIcon(Stop20Filled, Stop20Regular);
const CopyIcon = bundleIcon(Copy20Filled, Copy20Regular);
const DensityIcon = bundleIcon(TextDensity20Filled, TextDensity20Regular);

/** Every group the bar can show, in bar order */
const BAR_GROUPS = ["retrieve", "record", "state", "action", "data", "view"];

const useStyles = makeStyles({
	// The card around the bar gives the padding; the toolbar keeps a small inset
	bar: { paddingInline: tokens.spacingHorizontalXS },
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
	/**
	 * The entity's close actions (#149). When set, a Close menu takes Deactivate's place: one item
	 * per action, with the reason as a tooltip when one is off.
	 */
	closeActions?: CloseMenuItem[];
	/** Why the whole Close menu is off (e.g. a one-to-many query), if it is */
	closeDisabledReason?: string;
	/** Runs when a close action is chosen, with its key */
	onCloseAction?: (key: string) => void;
	/** Delete selected records (or all if none selected) */
	onDelete?: () => void;
	/** Bulk delete (creates async job) */
	onBulkDelete?: () => void;
	/** Whether delete is available (user has privilege) */
	canDelete?: boolean;
	/** Whether bulk delete is available (user has privilege) */
	canBulkDelete?: boolean;
	/** Why Delete is off for this query's shape (aggregate, or rows that repeat) */
	deleteDisabledReason?: string;
	/** Execute a specific workflow directly */
	onRunSpecificWorkflow?: (workflow: WorkflowInfo) => void;
	/** Whether workflow execution is available (user has privilege) */
	canRunWorkflow?: boolean;
	/** Why Run workflow is off for this query's shape */
	workflowDisabledReason?: string;
	/**
	 * Activate / Deactivate / Close show: the Write check finished and didn't say no. Without the
	 * privilege they're hidden (#232). Default true.
	 */
	canChangeState?: boolean;
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
	/** Fetch the on-demand workflows of an entity (the results' entity, #182) */
	onFetchWorkflows?: (entityLogicalName?: string) => Promise<WorkflowInfo[]>;
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
	/** Density of the grid and the query tree (#228) */
	density?: GridDensity;
	/** Change the density; the menu is hidden without it */
	onDensityChange?: (density: GridDensity) => void;
}

/** One item of the Close menu (#149) */
export interface CloseMenuItem {
	key: string;
	label: string;
	/** A divider goes above the first destructive item */
	destructive?: boolean;
}

export function ResultsCommandBar({
	selectedCount,
	onOpen,
	onCopyUrl,
	onActivate,
	onDeactivate,
	activateDisabledReason,
	deactivateDisabledReason,
	closeActions,
	closeDisabledReason,
	onCloseAction,
	onDelete,
	onBulkDelete,
	canDelete = false,
	canBulkDelete = false,
	deleteDisabledReason,
	onRunSpecificWorkflow,
	canRunWorkflow = false,
	workflowDisabledReason,
	canChangeState = true,
	recordActionsDisabledReason,
	deleteAllDisabledReason,
	onFetchWorkflows,
	onExport,
	onExportLocal,
	loadedCount = 0,
	onCopyRows,
	searchText = "",
	density = "medium",
	onDensityChange,
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
	const reasons = resultsBarReasons({
		selectedCount,
		loadedCount,
		recordActionsDisabledReason,
		deleteDisabledReason,
		workflowDisabledReason,
		deleteAllDisabledReason,
		canBulkDelete,
		activateDisabledReason,
		deactivateDisabledReason,
		closeDisabledReason,
	});
	const [editPanelOpen, setEditPanelOpen] = useState(false);
	const [addPanelOpen, setAddPanelOpen] = useState(false);
	// Copy menu: values as shown in the grid, or raw (#50)
	const [copyMode, setCopyMode] = useState<CopyValueMode>("shown");

	// Workflow menu: loaded when the menu opens, for the results' entity (#182)
	const {
		workflows,
		isLoading: isLoadingWorkflows,
		loaded: workflowsLoaded,
		load: handleWorkflowMenuOpen,
	} = useEntityWorkflows(entityName, onFetchWorkflows);

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
				const removed = columns
					.filter((col) => !updatedNames.has(col.name))
					.map((col) => col.name);
				if (removed.length > 0) onRemoveColumns(removed);
			}
			// Apply reordering
			onReorderColumns?.(updatedColumns);
			setEditPanelOpen(false);
		},
		[columns, onReorderColumns, onRemoveColumns],
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
		[onAddColumn, onAddRelatedColumns],
	);

	const handleOpenAddFromEdit = useCallback(() => {
		setAddPanelOpen(true);
	}, []);

	const commands: BarCommand[] = [];
	if (isRetrieveAllInProgress && onStopRetrieveAll) {
		commands.push({
			id: "stop",
			label: "Stop",
			icon: StopIcon,
			group: "retrieve",
			pinned: true,
			kind: "button",
			onClick: onStopRetrieveAll,
		});
	}
	commands.push(
		{
			id: "open",
			label: "Open",
			ariaLabel: "Open record",
			icon: OpenIcon,
			group: "record",
			priority: 50,
			kind: "button",
			disabledReason: reasons.open,
			onClick: onOpen,
		},
		{
			id: "copy-url",
			label: "Copy URL",
			ariaLabel: "Copy record URL",
			icon: LinkIcon,
			group: "record",
			priority: 40,
			kind: "button",
			disabledReason: reasons.copyUrl,
			onClick: onCopyUrl,
		},
	);

	// Activate / Deactivate (#143), or Close in Deactivate's place for close-message entities (#149).
	// Hidden without the Write privilege, and Close with no item the user can use (#232).
	if (canChangeState && onActivate) {
		commands.push({
			id: "activate",
			label: "Activate",
			ariaLabel: "Activate selected records",
			icon: ActivateIcon,
			group: "state",
			priority: 20,
			kind: "button",
			disabledReason: reasons.activate,
			onClick: onActivate,
		});
	}
	if (canChangeState && closeActions) {
		if (closeActions.length > 0 && onCloseAction) {
			commands.push({
				id: "close",
				label: "Close",
				ariaLabel: "Close selected records",
				icon: CloseIcon,
				group: "state",
				priority: 25,
				kind: "menu",
				disabledReason: reasons.close,
				menu: (
					<MenuList>
						<MenuGroup>
							{entityDisplayName && (
								<MenuGroupHeader>{entityDisplayName}</MenuGroupHeader>
							)}
							{closeActions.map((item, i) => (
								<Fragment key={item.key}>
									{i > 0 &&
										item.destructive &&
										!closeActions[i - 1].destructive && <MenuDivider />}
									<MenuItem
										onClick={() => onCloseAction(item.key)}
										disabled={!!reasons.close}
									>
										{item.label}...
									</MenuItem>
								</Fragment>
							))}
						</MenuGroup>
					</MenuList>
				),
			});
		}
	} else if (canChangeState && onDeactivate) {
		commands.push({
			id: "deactivate",
			label: "Deactivate",
			ariaLabel: "Deactivate selected records",
			icon: DeactivateIcon,
			group: "state",
			priority: 25,
			kind: "button",
			disabledReason: reasons.deactivate,
			onClick: onDeactivate,
		});
	}

	if (canDelete) {
		commands.push({
			id: "delete",
			label: hasSelection ? "Delete" : "Delete all",
			ariaLabel: hasSelection ? "Delete selected records" : "Delete all records from view",
			icon: DeleteIcon,
			group: "action",
			priority: 45,
			kind: "split",
			danger: true,
			disabledReason: reasons.delete,
			onClick: onDelete,
			// The menu lists Delete itself, so More shows it as it is
			menuHasPrimary: true,
			menu: (
				<MenuList>
					<MenuItem
						icon={<Delete20Regular />}
						onClick={onDelete}
						disabled={!!reasons.delete}
					>
						Delete{hasSelection ? ` (${selectedCount})` : " All"}
					</MenuItem>
					{canBulkDelete && (
						<MenuItem
							icon={<DocumentBulletList20Regular />}
							onClick={onBulkDelete}
							disabled={!!reasons.delete}
						>
							Bulk delete (job)
						</MenuItem>
					)}
				</MenuList>
			),
		});
	}

	if (canRunWorkflow) {
		commands.push({
			id: "run-workflow",
			label: "Run workflow",
			ariaLabel: "Run workflow on selected records",
			icon: WorkflowIcon,
			group: "action",
			priority: 22,
			kind: "menu",
			disabledReason: reasons.runWorkflow,
			// The results' entity's workflows load when the menu opens (#182)
			onMenuOpenChange: (open) => {
				if (open) handleWorkflowMenuOpen();
			},
			menu: (
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
			),
		});
	}

	commands.push(
		// Copy rows as tab-separated values with headers; pastes into Excel (#50)
		{
			id: "copy",
			label: "Copy",
			icon: CopyIcon,
			group: "data",
			priority: 30,
			kind: "menu",
			disabledReason: reasons.copy,
			menu: (
				<MenuList
					checkedValues={{ copyValues: [copyMode] }}
					onCheckedValueChange={(_, data) =>
						setCopyMode(data.checkedItems[0] as CopyValueMode)
					}
				>
					<MenuItem
						disabled={selectedCount === 0}
						onClick={() => onCopyRows?.("selected", copyMode)}
					>
						Selected rows ({selectedCount})
					</MenuItem>
					<MenuItem onClick={() => onCopyRows?.("all", copyMode)}>
						{searchText.trim() ? "All matching rows" : "All loaded rows"} ({loadedCount}
						)
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
			),
		},
		{
			id: "export",
			label: isExporting ? "Exporting..." : "Export to Excel",
			ariaLabel: "Export to Excel",
			icon: ExcelIcon,
			group: "data",
			priority: 15,
			kind: "menu",
			busy: isExporting,
			disabledReason: isExporting ? "Export in progress..." : undefined,
			menu: (
				<MenuList>
					<Tooltip
						content={
							canExport
								? "Export using Dataverse API (all records)"
								: exportDisabledReason ||
									"Save as a view first to enable server export"
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
			),
		},
		{
			id: "edit-columns",
			label: "Edit columns",
			icon: ColumnsIcon,
			group: "view",
			priority: 30,
			kind: "button",
			onClick: handleEditColumnsClick,
		},
	);

	// Density of the grid and the query tree, as in Dataverse Request Studio (#228); #291 moves it
	if (onDensityChange) {
		commands.push({
			id: "density",
			label: GRID_DENSITY_LABELS[density],
			icon: DensityIcon,
			group: "view",
			priority: 5,
			kind: "menu",
			hint: "Density (grid and tree)",
			menu: (
				<MenuList
					checkedValues={{ density: [density] }}
					onCheckedValueChange={(_, data) =>
						onDensityChange(data.checkedItems[0] as GridDensity)
					}
				>
					{(Object.keys(GRID_DENSITY_LABELS) as GridDensity[]).map((d) => (
						<MenuItemRadio key={d} name="density" value={d}>
							{GRID_DENSITY_LABELS[d]}
						</MenuItemRadio>
					))}
				</MenuList>
			),
		});
	}

	return (
		<>
			<OverflowToolbar
				ariaLabel="Record actions"
				commands={commands}
				// Stop, the state commands and Delete / Run workflow come and go with Retrieve all and
				// the privilege checks; declaring them keeps the bar (and an open menu) when they do
				groups={BAR_GROUPS}
				className={styles.bar}
			/>

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
