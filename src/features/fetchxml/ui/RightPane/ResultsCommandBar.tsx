/**
 * Command bar for DataGrid with row actions
 * Similar to Power Apps Model Driven Apps view command bar
 */

import { Fragment, useState, useCallback, type ReactElement, type ReactNode } from "react";
import {
	Button,
	Overflow,
	OverflowItem as FluentOverflowItem,
	OverflowDivider,
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
	useIsOverflowItemVisible,
	useOverflowMenu,
} from "@fluentui/react-components";
import {
	Open20Regular,
	Link20Regular,
	CheckmarkCircle20Regular,
	DismissCircle20Regular,
	CircleOff20Regular,
	Delete20Regular,
	ColumnTriple20Regular,
	Play20Regular,
	DocumentBulletList20Regular,
	Info20Regular,
	Stop20Regular,
	Copy20Regular,
	MoreHorizontal20Regular,
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
	// Wrapper: takes the card's free width and may shrink (as the FetchXML action bar's container does)
	bar: {
		flex: 1,
		minWidth: 0,
		display: "flex",
	},
	// The measured row: flat, shrinkable, clipped. Items that don't fit move into "More actions"
	primary: {
		display: "flex",
		alignItems: "center",
		gap: "2px",
		flex: 1,
		minWidth: 0,
		overflow: "hidden",
	},
	measuredItem: {
		display: "inline-flex",
		alignItems: "center",
		flexShrink: 0,
		whiteSpace: "nowrap",
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

/** One item of the Close menu (#149) */
export interface CloseMenuItem {
	key: string;
	label: string;
	/** A divider goes above the first destructive item */
	destructive?: boolean;
	/** Why this action is off (e.g. a missing privilege), if it is */
	disabledReason?: string;
}

/** Overflow must measure a real DOM element, even when an action is a Menu or a custom component. */
function OverflowItem({
	id,
	groupId,
	children,
}: {
	id: string;
	groupId: string;
	children: ReactNode;
}) {
	const styles = useStyles();
	return (
		<FluentOverflowItem id={id} groupId={groupId}>
			<span className={styles.measuredItem}>{children}</span>
		</FluentOverflowItem>
	);
}

/** A "More actions" entry; it is only listed while its toolbar item doesn't fit in the row */
function OverflowMenuEntry({ id, children }: { id: string; children: ReactNode }) {
	const visible = useIsOverflowItemVisible(id);
	return visible ? null : <>{children}</>;
}

/** The "More actions" button and menu; it appears only when some items are hidden in the row */
function OverflowMoreMenu({ children }: { children: ReactNode }) {
	const { ref, isOverflowing } = useOverflowMenu<HTMLButtonElement>();
	if (!isOverflowing) return null;
	return (
		<Menu positioning="below-end">
			<MenuTrigger disableButtonEnhancement>
				<Button
					ref={ref}
					appearance="subtle"
					size="small"
					icon={<MoreHorizontal20Regular />}
					aria-label="More actions"
				/>
			</MenuTrigger>
			<MenuPopover>
				<MenuList>{children}</MenuList>
			</MenuPopover>
		</Menu>
	);
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
		<ToolbarButton
			appearance="subtle"
			icon={icon}
			disabled={disabled}
			disabledFocusable={!!reason}
			onClick={onClick}
			aria-label={ariaLabel}
		>
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
	closeActions,
	closeDisabledReason,
	onCloseAction,
	onDelete,
	onBulkDelete,
	canDelete = false,
	canBulkDelete = false,
	deleteDisabled = false,
	onRunSpecificWorkflow,
	canRunWorkflow = false,
	workflowDisabled = false,
	recordActionsDisabledReason,
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
	const isDeleteDisabled = deleteDisabled || recordActionsUnavailable;
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

	// Menu contents, shared by the inline toolbar menus and their "More actions" submenus
	const closeMenuItems = () =>
		closeActions?.map((item, i) => {
			const menuItem = (
				<MenuItem
					disabled={!!item.disabledReason}
					disabledFocusable={!!item.disabledReason}
					onClick={() => onCloseAction?.(item.key)}
				>
					{item.label}...
				</MenuItem>
			);
			return (
				<Fragment key={item.key}>
					{i > 0 && item.destructive && !closeActions[i - 1].destructive && (
						<MenuDivider />
					)}
					{item.disabledReason ? (
						<Tooltip
							content={item.disabledReason}
							relationship="description"
							positioning="after"
							withArrow
						>
							{menuItem}
						</Tooltip>
					) : (
						menuItem
					)}
				</Fragment>
			);
		});

	const deleteMenuItems = () => (
		<>
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
		</>
	);

	const workflowMenuItems = () => (
		<>
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
		</>
	);

	const copyMenuItems = () => (
		<>
			<MenuItem
				disabled={selectedCount === 0}
				onClick={() => onCopyRows?.("selected", copyMode)}
			>
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
		</>
	);

	const exportMenuItems = () => (
		<>
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
		</>
	);

	/** A "More actions" entry that opens a submenu with the same content as the inline menu */
	const submenuEntry = (
		label: string,
		icon: ReactElement,
		disabled: boolean,
		content: () => ReactNode,
		onOpen?: () => void,
	) => (
		<Menu positioning="after-top" onOpenChange={(_, data) => data.open && onOpen?.()}>
			<MenuTrigger disableButtonEnhancement>
				<MenuItem hasSubmenu icon={icon} disabled={disabled}>
					{label}
				</MenuItem>
			</MenuTrigger>
			<MenuPopover>
				<MenuList>{content()}</MenuList>
			</MenuPopover>
		</Menu>
	);

	return (
		<>
			{/* Same structure as the FetchXML action bar (OverflowActionBar): the flat row is the measured element */}
			<div className={styles.bar}>
				<Overflow>
					<div className={styles.primary} role="toolbar" aria-label="Record actions">
						{isRetrieveAllInProgress && onStopRetrieveAll && (
							<>
								<OverflowItem id="stop" groupId="record">
									<ToolbarButton
										appearance="subtle"
										icon={<Stop20Regular />}
										onClick={onStopRetrieveAll}
									>
										Stop
									</ToolbarButton>
								</OverflowItem>
								<OverflowDivider groupId="record">
									<ToolbarDivider />
								</OverflowDivider>
							</>
						)}
						<OverflowItem id="open" groupId="record">
							<ToolbarButton
								appearance="subtle"
								icon={<Open20Regular />}
								disabled={!singleSelection || recordActionsUnavailable}
								onClick={onOpen}
								aria-label="Open record"
							>
								Open
							</ToolbarButton>
						</OverflowItem>

						<OverflowItem id="copyUrl" groupId="record">
							<ToolbarButton
								appearance="subtle"
								icon={<Link20Regular />}
								disabled={!singleSelection || recordActionsUnavailable}
								onClick={onCopyUrl}
								aria-label="Copy record URL"
							>
								Copy URL
							</ToolbarButton>
						</OverflowItem>

						{recordActionsDisabledReason && (
							<OverflowItem id="recordInfo" groupId="record">
								<Tooltip
									content={recordActionsDisabledReason}
									relationship="description"
									withArrow
								>
									<ToolbarButton
										appearance="transparent"
										icon={<Info20Regular />}
									>
										Record actions unavailable
									</ToolbarButton>
								</Tooltip>
							</OverflowItem>
						)}

						<OverflowDivider groupId="state">
							<ToolbarDivider />
						</OverflowDivider>

						{/* Activate / Deactivate (#143); the tooltip says why one is off for this entity */}
						<OverflowItem id="activate" groupId="state">
							<StateButton
								label="Activate"
								ariaLabel="Activate selected records"
								icon={<CheckmarkCircle20Regular />}
								onClick={onActivate}
								disabled={
									stateButtonsDisabled || !onActivate || !!activateDisabledReason
								}
								reason={activateDisabledReason}
							/>
						</OverflowItem>

						{closeActions && closeActions.length > 0 && closeDisabledReason ? (
							<OverflowItem id="close" groupId="state">
								<StateButton
									label="Close"
									ariaLabel="Close selected records"
									icon={<CircleOff20Regular />}
									disabled
									reason={closeDisabledReason}
								/>
							</OverflowItem>
						) : closeActions && closeActions.length > 0 ? (
							// Close (#149): the entity's close messages, in place of Deactivate
							<OverflowItem id="close" groupId="state">
								<Menu positioning="below-start">
									<MenuTrigger disableButtonEnhancement>
										<MenuButton
											appearance="subtle"
											icon={<CircleOff20Regular />}
											disabled={stateButtonsDisabled || !onCloseAction}
											aria-label="Close selected records"
										>
											Close
										</MenuButton>
									</MenuTrigger>
									<MenuPopover>
										<MenuList>
											<MenuGroup>
												{entityDisplayName && (
													<MenuGroupHeader>
														{entityDisplayName}
													</MenuGroupHeader>
												)}
												{closeMenuItems()}
											</MenuGroup>
										</MenuList>
									</MenuPopover>
								</Menu>
							</OverflowItem>
						) : (
							<OverflowItem id="deactivate" groupId="state">
								<StateButton
									label="Deactivate"
									ariaLabel="Deactivate selected records"
									icon={<DismissCircle20Regular />}
									onClick={onDeactivate}
									disabled={
										stateButtonsDisabled ||
										!onDeactivate ||
										!!deactivateDisabledReason
									}
									reason={deactivateDisabledReason}
								/>
							</OverflowItem>
						)}

						<OverflowDivider groupId="delete">
							<ToolbarDivider />
						</OverflowDivider>

						{canDelete && (
							<OverflowItem id="delete" groupId="delete">
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
													hasSelection
														? "Delete selected records"
														: "Delete all records from view"
												}
											>
												{hasSelection ? "Delete" : "Delete all"}
											</SplitButton>
										)}
									</MenuTrigger>
									<MenuPopover>
										<MenuList>{deleteMenuItems()}</MenuList>
									</MenuPopover>
								</Menu>
							</OverflowItem>
						)}

						{canRunWorkflow && (
							<OverflowItem id="workflow" groupId="workflow">
								<Menu
									positioning="below-end"
									onOpenChange={(_, data) =>
										data.open && handleWorkflowMenuOpen()
									}
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
										<MenuList>{workflowMenuItems()}</MenuList>
									</MenuPopover>
								</Menu>
							</OverflowItem>
						)}

						<OverflowDivider groupId="copy">
							<ToolbarDivider />
						</OverflowDivider>

						{/* Copy rows as tab-separated values with headers; pastes into Excel (#50) */}
						<OverflowItem id="copy" groupId="copy">
							<Menu
								checkedValues={{ copyValues: [copyMode] }}
								onCheckedValueChange={(_, data) =>
									setCopyMode(data.checkedItems[0] as CopyValueMode)
								}
							>
								<MenuTrigger disableButtonEnhancement>
									<MenuButton
										appearance="subtle"
										icon={<Copy20Regular />}
										disabled={loadedCount === 0}
									>
										Copy
									</MenuButton>
								</MenuTrigger>
								<MenuPopover>
									<MenuList>{copyMenuItems()}</MenuList>
								</MenuPopover>
							</Menu>
						</OverflowItem>

						{/* Export Menu Button */}
						<OverflowItem id="export" groupId="copy">
							<Menu>
								<MenuTrigger disableButtonEnhancement>
									<Tooltip
										content={
											isExporting
												? "Export in progress..."
												: "Export to Excel"
										}
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
									<MenuList>{exportMenuItems()}</MenuList>
								</MenuPopover>
							</Menu>
						</OverflowItem>

						<OverflowDivider groupId="editColumns">
							<ToolbarDivider />
						</OverflowDivider>

						<OverflowItem id="editColumns" groupId="editColumns">
							<ToolbarButton
								appearance="subtle"
								icon={<ColumnTriple20Regular />}
								onClick={handleEditColumnsClick}
								aria-label="Edit columns"
							>
								Edit columns
							</ToolbarButton>
						</OverflowItem>
						<OverflowMoreMenu>
							<OverflowMenuEntry id="stop">
								{isRetrieveAllInProgress && onStopRetrieveAll && (
									<MenuItem icon={<Stop20Regular />} onClick={onStopRetrieveAll}>
										Stop
									</MenuItem>
								)}
							</OverflowMenuEntry>
							<OverflowMenuEntry id="open">
								<MenuItem
									icon={<Open20Regular />}
									disabled={!singleSelection || recordActionsUnavailable}
									onClick={onOpen}
								>
									Open
								</MenuItem>
							</OverflowMenuEntry>
							<OverflowMenuEntry id="copyUrl">
								<MenuItem
									icon={<Link20Regular />}
									disabled={!singleSelection || recordActionsUnavailable}
									onClick={onCopyUrl}
								>
									Copy URL
								</MenuItem>
							</OverflowMenuEntry>
							<OverflowMenuEntry id="recordInfo">
								{recordActionsDisabledReason && (
									<MenuItem icon={<Info20Regular />} disabled>
										Record actions unavailable
									</MenuItem>
								)}
							</OverflowMenuEntry>
							<OverflowMenuEntry id="activate">
								<MenuItem
									icon={<CheckmarkCircle20Regular />}
									disabled={
										stateButtonsDisabled ||
										!onActivate ||
										!!activateDisabledReason
									}
									onClick={onActivate}
								>
									Activate
								</MenuItem>
							</OverflowMenuEntry>
							<OverflowMenuEntry id="close">
								{closeActions && closeActions.length > 0 && closeDisabledReason && (
									<MenuItem icon={<CircleOff20Regular />} disabled>
										Close
									</MenuItem>
								)}
								{closeActions &&
									closeActions.length > 0 &&
									!closeDisabledReason &&
									submenuEntry(
										"Close",
										<CircleOff20Regular />,
										stateButtonsDisabled || !onCloseAction,
										closeMenuItems,
									)}
							</OverflowMenuEntry>
							{!closeActions?.length && (
								<OverflowMenuEntry id="deactivate">
									<MenuItem
										icon={<DismissCircle20Regular />}
										disabled={
											stateButtonsDisabled ||
											!onDeactivate ||
											!!deactivateDisabledReason
										}
										onClick={onDeactivate}
									>
										Deactivate
									</MenuItem>
								</OverflowMenuEntry>
							)}
							<OverflowMenuEntry id="delete">
								{canDelete &&
									submenuEntry(
										"Delete",
										<Delete20Regular />,
										isDeleteDisabled,
										deleteMenuItems,
									)}
							</OverflowMenuEntry>
							<OverflowMenuEntry id="workflow">
								{canRunWorkflow &&
									submenuEntry(
										"Run workflow",
										<Play20Regular />,
										!hasSelection || isWorkflowDisabled,
										workflowMenuItems,
										handleWorkflowMenuOpen,
									)}
							</OverflowMenuEntry>
							<OverflowMenuEntry id="copy">
								{submenuEntry(
									"Copy",
									<Copy20Regular />,
									loadedCount === 0,
									copyMenuItems,
								)}
							</OverflowMenuEntry>
							<OverflowMenuEntry id="export">
								{submenuEntry(
									"Export to Excel",
									<ExcelIcon />,
									isExporting,
									exportMenuItems,
								)}
							</OverflowMenuEntry>
							<OverflowMenuEntry id="editColumns">
								<MenuItem
									icon={<ColumnTriple20Regular />}
									onClick={handleEditColumnsClick}
								>
									Edit columns
								</MenuItem>
							</OverflowMenuEntry>
						</OverflowMoreMenu>
					</div>
				</Overflow>
			</div>

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
