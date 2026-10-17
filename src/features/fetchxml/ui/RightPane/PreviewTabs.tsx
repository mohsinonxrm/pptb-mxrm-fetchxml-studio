/**
 * Tabbed preview panel with FetchXML editor and Results grid
 */

import { useState, useCallback, useMemo, useRef, type ReactNode } from "react";
import {
	TabList,
	Tab,
	type SelectTabData,
	type SelectTabEvent,
	Button,
	makeStyles,
	Toolbar,
	ToolbarDivider,
	ProgressBar,
	Spinner,
	Tooltip,
	tokens,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	MessageBarActions,
} from "@fluentui/react-components";
import {
	Play24Regular,
	Dismiss16Regular,
	Settings20Regular,
	ArrowReply24Regular,
	Code20Regular,
} from "@fluentui/react-icons";
import { returnFetchXmlToInvokingTool } from "../../api/invocation";
import { FetchXmlEditor } from "./FetchXmlEditor";
import { CodePanel } from "./CodePanel";
import { LayoutXmlViewer } from "./LayoutXmlViewer";
import { ResultsGrid, type QueryResult, type SortChangeData } from "./ResultsGrid";
import { ResultsCommandBar } from "./ResultsCommandBar";
import type { RelatedEntityColumn } from "./AddColumnsPanel";
import type { PagingNotice } from "../../api/paging";
import { isAttributeInScope, useAdvancedFindOnly } from "../../state/metadataScope";
import { NON_RETURNING_LINK_TYPES } from "../../model/fetchxmlRules";
import type { AttributeMetadata, RelationshipMetadata } from "../../api/pptbClient";
import type { FetchNode } from "../../model/nodes";
import type { ParseResult, ParseWarning } from "../../model/fetchxmlParser";
import {
	validateFetchXmlSyntax,
	parseFetchXml,
	getBlockingWarnings,
	getExecutionErrors,
	getParseToTreeChanges,
} from "../../model/fetchxmlParser";
import { generateFetchXml } from "../../model/fetchxml";
import type { LoadFetchXmlOptions } from "../../state/builderContext";
import { ParseChangesDialog } from "./ParseChangesDialog";
import {
	EditorSyncContext,
	type EditorSyncAction,
	type EditorSyncResult,
} from "./editorSync";
import { ParseIssueList, type ParseIssueItem } from "./ParseIssueList";
import type { LayoutXmlConfig } from "../../model/layoutxml";
import type { WorkflowInfo } from "../../api/pptbClient";
import type { DisplaySettings } from "../../model/displaySettings";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		overflow: "hidden",
	},
	header: {
		display: "flex",
		alignItems: "center",
		justifyContent: "space-between",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		paddingRight: "16px",
	},
	progressContainer: {
		paddingLeft: "16px",
		paddingRight: "16px",
		paddingTop: "8px",
		paddingBottom: "8px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
	},
	tabContent: {
		flex: 1,
		overflow: "hidden",
		padding: "6px",
		backgroundColor: tokens.colorNeutralBackground3, // Neutral canvas background
	},
	// Shared surface card styling (floating cards with shadow)
	surfaceCard: {
		borderRadius: tokens.borderRadiusMedium,
		boxShadow: tokens.shadow16,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
	},
	// Toolbar card styling
	toolbarCard: {
		borderRadius: tokens.borderRadiusMedium,
		boxShadow: tokens.shadow16,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		padding: "4px 8px",
		flexShrink: 0,
	},
	// Grid card styling
	gridCard: {
		borderRadius: tokens.borderRadiusMedium,
		boxShadow: tokens.shadow16,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		display: "flex",
		flexDirection: "column",
		flex: 1,
		minHeight: "240px",
		minWidth: "480px",
		overflow: "hidden",
		marginTop: "8px",
	},
	// Code card for FetchXML tab
	codeCard: {
		borderRadius: tokens.borderRadiusMedium,
		boxShadow: tokens.shadow16,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		height: "100%",
		display: "flex",
		flexDirection: "column",
	},
	// Results tab layout grid
	resultsLayout: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		overflow: "hidden",
	},
	messageBarContainer: {
		paddingLeft: "16px",
		paddingRight: "16px",
		paddingTop: "8px",
		paddingBottom: "8px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
	},
});

/**
 * Returns the current FetchXML to the tool that launched this window (T2T callee).
 *
 * This is distinct from the host-injected "Return to [Caller]" banner: that banner
 * is a plain "go back" affordance and resolves the caller's promise with null. Only
 * an explicit returnData() call carries the query back — which is what this does.
 */
function ReturnFetchXmlButton({ getCurrentXml }: { getCurrentXml: () => string }) {
	const [isReturning, setIsReturning] = useState(false);

	const handleReturn = useCallback(async () => {
		setIsReturning(true);
		try {
			await returnFetchXmlToInvokingTool(getCurrentXml());
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			await window.toolboxAPI?.utils?.showNotification?.({
				title: "Cannot return FetchXML",
				body: message,
				type: "error",
			});
		} finally {
			setIsReturning(false);
		}
	}, [getCurrentXml]);

	return (
		<Tooltip
			content="Return the current FetchXML to the tool that launched this window"
			relationship="description"
		>
			<Button
				appearance="subtle"
				icon={isReturning ? <Spinner size="tiny" /> : <ArrowReply24Regular />}
				onClick={handleReturn}
				disabled={isReturning}
			>
				Return FetchXML
			</Button>
		</Tooltip>
	);
}

interface PreviewTabsProps {
	xml: string;
	/** LayoutXML string for the grid column configuration */
	layoutXml?: string;
	result: QueryResult | null;
	isExecuting?: boolean;
	/** Whether more pages are being loaded (for progress display) */
	isLoadingMore?: boolean;
	/**
	 * Execute the tree's query. After an editor-mode sync (#102) the new tree and its XML are
	 * passed along, because builder state updates arrive on the next render.
	 */
	onExecute?: (synced?: { xml: string; fetchQuery: FetchNode }) => void;
	/** Export via Dataverse ExportToExcel API */
	onExport?: () => void;
	/** Export locally using exceljs */
	onExportLocal?: () => void;
	onParseToTree?: (xmlString: string, options?: LoadFetchXmlOptions) => ParseResult;
	/** Multi-entity attribute metadata: Map<entityLogicalName, Map<attributeLogicalName, AttributeMetadata>> */
	attributeMetadata?: Map<string, Map<string, AttributeMetadata>>;
	fetchQuery?: FetchNode | null;
	/** Error message from a failed query execution (e.g. Dataverse API error) */
	executeError?: string;
	/** Callback to dismiss the execute error */
	onDismissExecuteError?: () => void;
	/** Column layout configuration for ordering and sizing */
	columnConfig?: LayoutXmlConfig | null;
	/** Callback when column width changes */
	onColumnResize?: (columnName: string, width: number) => void;
	/** Callback when columns are reordered */
	onReorderColumns?: (columns: LayoutXmlConfig["columns"]) => void;
	/** Callback when user wants to add a column from available attributes */
	onAddColumn?: (attributeName: string) => void;
	/** Callback when user wants to add columns from related entity */
	onAddRelatedColumns?: (columns: RelatedEntityColumn[]) => void;
	/** Callback when user removes a column */
	onRemoveColumns?: (columnKeys: string[]) => void;
	/** Callback when user changes sort on a column */
	onSortChange?: (data: SortChangeData) => void;
	/** Optional SaveViewButton to render in the toolbar */
	saveViewButton?: ReactNode;
	/** Optional SendToToolButton to render in the toolbar */
	sendToToolButton?: ReactNode;
	/** True when this tool was launched by another tool via T2T — shows the Return button */
	isCallee?: boolean;
	/** Callback when user scrolls near bottom (infinite scroll) */
	onLoadMore?: () => void;
	/** Paging notice above the grid: failed page, Stop, simple paging (#60) */
	pagingNotice?: PagingNotice | null;
	isRetrieveAllInProgress?: boolean;
	onStopRetrieveAll?: () => void;
	onRetryPaging?: () => void;
	onDismissPagingNotice?: () => void;
	/** Whether export is available (requires a saved view) */
	canExport?: boolean;
	/** Whether export is in progress */
	isExporting?: boolean;
	/** Export error message */
	exportError?: string;
	/** Callback to dismiss export error */
	onDismissExportError?: () => void;
	/** Tooltip text for disabled export button */
	exportDisabledReason?: string;
	/** Entity display name */
	entityDisplayName?: string;
	/** Lookup relationships for related columns */
	lookupRelationships?: RelationshipMetadata[];
	/** One-to-many relationships for 1-N related columns */
	oneToManyRelationships?: RelationshipMetadata[];
	/** Whether relationship data is loading */
	isLoadingRelationships?: boolean;
	/** Callback to load attributes for a related entity */
	onLoadRelatedAttributes?: (entityLogicalName: string) => Promise<AttributeMetadata[]>;
	/** Callback to reset columns to default */
	onResetToDefault?: () => void;
	// Record action callbacks
	/** Callback when user clicks Open (opens record(s) in new tab) */
	onOpenRecord?: (recordIds: string[]) => void;
	/** Callback when user clicks Copy URL (copies record URL(s)) */
	onCopyRecordUrl?: (recordIds: string[]) => void;
	/** Callback when user clicks Delete */
	onDeleteRecords?: (recordIds: string[]) => void;
	/** Callback when user clicks Bulk Delete */
	onBulkDeleteRecords?: (recordIds: string[]) => void;
	/** Callback when user clicks Activate */
	onActivateRecords?: (recordIds: string[]) => void;
	/** Callback when user clicks Deactivate */
	onDeactivateRecords?: (recordIds: string[]) => void;
	/** Callback when user wants to run a specific workflow directly */
	onRunSpecificWorkflow?: (workflow: WorkflowInfo, recordIds: string[]) => void;
	/** Whether user can delete records */
	canDelete?: boolean;
	/** Whether user can bulk delete records */
	canBulkDelete?: boolean;
	/** Whether user can run workflows */
	canRunWorkflow?: boolean;
	/** Fetch available workflows for the entity */
	onFetchWorkflows?: () => Promise<WorkflowInfo[]>;
	/** Whether the current query is an aggregate query (disables delete/workflow) */
	isAggregateQuery?: boolean;
	/** Whether the query has 1-N or N-N relationships that cause row duplication */
	hasOneToManyRelationship?: boolean;
	/** When set, results have no record IDs and all record actions are disabled with this explanation */
	recordActionsDisabledReason?: string;
	/** Callback when selection changes in ResultsGrid */
	onSelectionChange?: (recordIds: string[]) => void;
	/** Callback to get currently selected record IDs */
	getSelectedRecordIds?: () => string[];
	/** Callback when user clicks Settings button */
	onOpenSettings?: () => void;
	/** Display settings (logical names, value format) */
	displaySettings?: DisplaySettings;
	/** Primary key attribute from Dataverse entity metadata (e.g. "activityid" for activity entities) */
	primaryIdAttribute?: string;
}

export function PreviewTabs({
	xml,
	layoutXml,
	result,
	isExecuting,
	isLoadingMore,
	onExecute,
	onExport,
	onExportLocal,
	onParseToTree,
	attributeMetadata,
	fetchQuery,
	columnConfig,
	onColumnResize,
	onReorderColumns,
	onAddColumn,
	onAddRelatedColumns,
	onRemoveColumns,
	onSortChange,
	saveViewButton,
	sendToToolButton,
	isCallee,
	onLoadMore,
	pagingNotice,
	isRetrieveAllInProgress,
	onStopRetrieveAll,
	onRetryPaging,
	onDismissPagingNotice,
	canExport,
	isExporting,
	exportError,
	onDismissExportError,
	exportDisabledReason,
	executeError,
	onDismissExecuteError,
	entityDisplayName,
	lookupRelationships,
	oneToManyRelationships,
	isLoadingRelationships,
	onLoadRelatedAttributes,
	onResetToDefault,
	onOpenRecord,
	onCopyRecordUrl,
	onDeleteRecords,
	onBulkDeleteRecords,
	onActivateRecords,
	onDeactivateRecords,
	onRunSpecificWorkflow,
	canDelete,
	canBulkDelete,
	canRunWorkflow,
	onFetchWorkflows,
	isAggregateQuery,
	hasOneToManyRelationship,
	recordActionsDisabledReason,
	onSelectionChange,
	getSelectedRecordIds,
	onOpenSettings,
	displaySettings,
	primaryIdAttribute,
}: PreviewTabsProps) {
	const styles = useStyles();
	const advancedFindOnly = useAdvancedFindOnly();
	const [selectedTab, setSelectedTab] = useState<"xml" | "layout" | "results" | "code">("xml");
	const [toolbarSelectedCount, setToolbarSelectedCount] = useState(0);
	const [selectedRecordIds, setSelectedRecordIds] = useState<string[]>([]);

	// Track editor mode state via ref (avoids re-renders on each keystroke)
	// isEditorModeActive in state controls the disabled prop + renders; the ref holds the latest XML.
	const [isEditorModeActive, setIsEditorModeActive] = useState(false);
	const editorXmlRef = useRef<string>("");
	const [editorValidationError, setEditorValidationError] = useState<ParseIssueItem[] | null>(
		null,
	);
	// Editor XML Dataverse might reject, awaiting "Execute anyway" (#89)
	const [pendingExecute, setPendingExecute] = useState<{
		xml: string;
		warnings: ParseWarning[];
	} | null>(null);
	// Related columns already returned by root link-entities Add columns would reuse (#114)
	const existingRelatedColumns = useMemo(
		() =>
			new Set(
				(fetchQuery?.entity.links ?? [])
					.filter((link) => !NON_RETURNING_LINK_TYPES.has(link.linkType))
					.flatMap((link) => link.attributes.map((a) => `${link.name}|${link.from}|${link.to}|${a.name}`)),
			),
		[fetchQuery],
	);

	// Whether the editor buffer has any content (Execute is enabled on it in editor mode, #102)
	const [editorHasContent, setEditorHasContent] = useState(false);
	// Cleaned FetchXML to show in the editor after a sync changed the query (#102)
	const [editorReplace, setEditorReplace] = useState<{ xml: string; id: number } | null>(null);
	// Editor sync awaiting confirmation because loading into the tree would change the query
	const [pendingSync, setPendingSync] = useState<{
		action: EditorSyncAction;
		changes: ParseWarning[];
		resolve: (confirmed: boolean) => void;
	} | null>(null);

	// The XML currently shown to the user: the live editor buffer when in editor
	// mode, otherwise the tree-generated XML. Used by the T2T "Return FetchXML" button.
	const getCurrentXml = useCallback(() => {
		return isEditorModeActive ? editorXmlRef.current || xml : xml;
	}, [isEditorModeActive, xml]);

	const handleEditorStateChange = useCallback((isActive: boolean, editorXml: string) => {
		editorXmlRef.current = editorXml;
		setIsEditorModeActive((prev) => (prev !== isActive ? isActive : prev));
		setEditorHasContent(isActive && editorXml.trim().length > 0);
		// A pending "Execute anyway" only applies to the exact XML it warned about
		setPendingExecute((prev) => (prev && (!isActive || prev.xml !== editorXml) ? null : prev));
		if (!isActive) {
			setEditorValidationError(null);
		}
	}, []);

	const handleTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
		setSelectedTab(data.value as "xml" | "layout" | "results" | "code");
		// Clear the editor validation error when returning to the XML tab so it
		// doesn't linger as a banner on top of the editor.
		if (data.value === "xml") {
			setEditorValidationError(null);
		}
	};

	/**
	 * Editor mode (#102): load the editor's FetchXML into the tree before an action, so the tree is
	 * always what runs, is saved and is sent. Invalid FetchXML is shown and stops the action; if
	 * loading would change the query, the user confirms the changes first (#89). When the editor
	 * already matches the tree, nothing is reloaded.
	 */
	const syncEditorToTree = useCallback(
		async (action: EditorSyncAction): Promise<EditorSyncResult> => {
			if (!isEditorModeActive) return { status: "unchanged" };
			const editorXml = editorXmlRef.current;

			const syntaxResult = validateFetchXmlSyntax(editorXml);
			if (!syntaxResult.valid) {
				setEditorValidationError([{ message: syntaxResult.error || "Invalid XML syntax" }]);
				return { status: "cancelled" };
			}
			const preview = parseFetchXml(editorXml);
			if (!preview.success || !preview.fetchNode) {
				setEditorValidationError(preview.errors);
				return { status: "cancelled" };
			}
			// Known to be rejected by Dataverse: block (Parse to Tree can often correct these)
			const executionErrors = getExecutionErrors(preview);
			if (executionErrors.length > 0) {
				setEditorValidationError(executionErrors);
				return { status: "cancelled" };
			}
			if (xml && generateFetchXml(preview.fetchNode) === xml) return { status: "unchanged" };

			const changes = getParseToTreeChanges(preview);
			if (changes.length > 0) {
				const confirmed = await new Promise<boolean>((resolve) =>
					setPendingSync({ action, changes, resolve }),
				);
				if (!confirmed) return { status: "cancelled" };
			}

			const result = onParseToTree?.(editorXml, { preserveLayout: true });
			if (!result?.success || !result.fetchNode) {
				setEditorValidationError(result?.errors ?? [{ message: "The FetchXML couldn't be loaded into the tree." }]);
				return { status: "cancelled" };
			}
			const syncedXml = generateFetchXml(result.fetchNode);
			// The tree changed the query (the user confirmed it): show what actually runs in the editor
			if (changes.length > 0) {
				editorXmlRef.current = syncedXml;
				setEditorReplace({ xml: syncedXml, id: Date.now() });
			}
			return { status: "synced", fetchXml: syncedXml, fetchQuery: result.fetchNode };
		},
		[isEditorModeActive, xml, onParseToTree],
	);

	const editorSyncValue = useMemo(
		() => ({ sync: syncEditorToTree, editorHasContent: isEditorModeActive && editorHasContent }),
		[syncEditorToTree, isEditorModeActive, editorHasContent],
	);

	const resolvePendingSync = (confirmed: boolean) => {
		pendingSync?.resolve(confirmed);
		setPendingSync(null);
	};

	/** Sync the editor into the tree (editor mode), then run the tree's query. */
	const runExecute = async () => {
		const synced = await syncEditorToTree("execute");
		if (synced.status === "cancelled") return;
		setSelectedTab("results");
		onExecute?.(synced.status === "synced" ? { xml: synced.fetchXml, fetchQuery: synced.fetchQuery } : undefined);
	};

	/** Run the editor XML despite warnings; Dataverse's own error is shown if it fails. */
	const handleExecuteAnyway = () => {
		if (!pendingExecute) return;
		setPendingExecute(null);
		void runExecute();
	};

	const handleExecute = () => {
		setEditorValidationError(null);
		setPendingExecute(null);

		if (isEditorModeActive) {
			// Can't be determined client-side: ask before sending. Errors are checked by the sync.
			const editorXml = editorXmlRef.current;
			const preview = parseFetchXml(editorXml);
			if (preview.success && getExecutionErrors(preview).length === 0) {
				const blocking = getBlockingWarnings(preview);
				if (blocking.length > 0) {
					setPendingExecute({ xml: editorXml, warnings: blocking });
					return;
				}
			}
		}
		void runExecute();
	};

	const handleSelectionChange = useCallback(
		(recordIds: string[]) => {
			setSelectedRecordIds(recordIds);
			// Also notify parent
			onSelectionChange?.(recordIds);
		},
		[onSelectionChange],
	);

	// If parent provided a getter, use ours; otherwise use internal state
	const getSelectedIds = useCallback(() => {
		return getSelectedRecordIds ? getSelectedRecordIds() : selectedRecordIds;
	}, [getSelectedRecordIds, selectedRecordIds]);

	return (
		<EditorSyncContext.Provider value={editorSyncValue}>
			<div className={styles.container}>
				<ParseChangesDialog
					action={pendingSync?.action ?? null}
					changes={pendingSync?.changes ?? []}
					onConfirm={() => resolvePendingSync(true)}
					onCancel={() => resolvePendingSync(false)}
				/>
				<div className={styles.header}>
					<TabList selectedValue={selectedTab} onTabSelect={handleTabSelect}>
						<Tab value="xml">FetchXML</Tab>
						<Tab value="layout">LayoutXML</Tab>
						<Tab value="results">Results</Tab>
						<Tab value="code" icon={<Code20Regular />}>
							Code
						</Tab>
					</TabList>
					<Toolbar size="small">
						<Button
							appearance="primary"
							icon={<Play24Regular />}
							onClick={handleExecute}
							disabled={isExecuting || isLoadingMore || (isEditorModeActive ? !editorHasContent : !xml)}
						>
							{isExecuting ? "Executing..." : "Execute"}
						</Button>
						{isCallee && (
							<>
								<ToolbarDivider />
								<ReturnFetchXmlButton getCurrentXml={getCurrentXml} />
							</>
						)}
						{selectedTab === "xml" && saveViewButton && (
							<>
								<ToolbarDivider />
								{saveViewButton}
							</>
						)}
						{sendToToolButton && (
							<>
								<ToolbarDivider />
								{sendToToolButton}
							</>
						)}
						<ToolbarDivider />
						<Button
							appearance="subtle"
							icon={<Settings20Regular />}
							onClick={onOpenSettings}
							aria-label="Settings"
						/>
					</Toolbar>
				</div>
				{(isExecuting || isLoadingMore) && (
					<div className={styles.progressContainer}>
						<ProgressBar />
					</div>
				)}
				{/* Export status MessageBar */}
				{isExporting && (
					<div className={styles.messageBarContainer}>
						<MessageBar intent="info">
							<MessageBarBody>
								<MessageBarTitle>Exporting</MessageBarTitle>
								Exporting view to Excel...
							</MessageBarBody>
						</MessageBar>
					</div>
				)}
				{exportError && !isExporting && (
					<div className={styles.messageBarContainer}>
						<MessageBar intent="error">
							<MessageBarBody>
								<MessageBarTitle>Export Failed</MessageBarTitle>
								{exportError}
							</MessageBarBody>
							<MessageBarActions
								containerAction={
									<Button
										appearance="transparent"
										icon={<Dismiss16Regular />}
										onClick={onDismissExportError}
										aria-label="Dismiss"
									/>
								}
							/>
						</MessageBar>
					</div>
				)}
				{pendingExecute && (
					<div className={styles.messageBarContainer}>
						<MessageBar intent="warning" layout="multiline">
							<MessageBarBody>
								<MessageBarTitle>Dataverse might reject this query</MessageBarTitle>
								<ParseIssueList items={pendingExecute.warnings} />
							</MessageBarBody>
							<MessageBarActions
								containerAction={
									<Button
										appearance="transparent"
										icon={<Dismiss16Regular />}
										onClick={() => setPendingExecute(null)}
										aria-label="Dismiss"
									/>
								}
							>
								<Button onClick={handleExecuteAnyway}>Execute anyway</Button>
							</MessageBarActions>
						</MessageBar>
					</div>
				)}
				{editorValidationError && (
					<div className={styles.messageBarContainer}>
						<MessageBar intent="error" layout="multiline">
							<MessageBarBody>
								<MessageBarTitle>Invalid FetchXML</MessageBarTitle>
								<ParseIssueList items={editorValidationError} />
							</MessageBarBody>
							<MessageBarActions
								containerAction={
									<Button
										appearance="transparent"
										icon={<Dismiss16Regular />}
										onClick={() => setEditorValidationError(null)}
										aria-label="Dismiss"
									/>
								}
							/>
						</MessageBar>
					</div>
				)}
				{executeError && (
					<div className={styles.messageBarContainer}>
						<MessageBar intent="error">
							<MessageBarBody>
								<MessageBarTitle>Execution Failed</MessageBarTitle>
								{executeError}
							</MessageBarBody>
							<MessageBarActions
								containerAction={
									<Button
										appearance="transparent"
										icon={<Dismiss16Regular />}
										onClick={onDismissExecuteError}
										aria-label="Dismiss"
									/>
								}
							/>
						</MessageBar>
					</div>
				)}
				<div className={styles.tabContent}>
					{/* FetchXML tab — always mounted so Monaco editor state (editor mode, content,
					    cursor position) survives tab switches. Hidden via display:none when inactive;
					    Monaco's automaticLayout re-layouts automatically when it becomes visible. */}
					<div
						className={styles.codeCard}
						style={selectedTab !== "xml" ? { display: "none" } : undefined}
					>
						<FetchXmlEditor
							xml={xml}
							onParseToTree={onParseToTree}
							onEditorStateChange={handleEditorStateChange}
							replaceRequest={editorReplace}
						/>
					</div>
					{selectedTab === "layout" && (
						<div className={styles.codeCard}>
							<LayoutXmlViewer layoutXml={layoutXml || ""} />
						</div>
					)}
					{selectedTab === "code" && (
						<div className={styles.codeCard}>
							<CodePanel fetchXml={xml} />
						</div>
					)}
					{selectedTab === "results" && (
						<div className={styles.resultsLayout}>
							<div className={styles.toolbarCard}>
								<ResultsCommandBar
									selectedCount={toolbarSelectedCount}
									onOpen={() => {
										const ids = getSelectedIds();
										if (ids?.length && onOpenRecord) {
											onOpenRecord(ids);
										}
									}}
									onCopyUrl={() => {
										const ids = getSelectedIds();
										if (ids?.length && onCopyRecordUrl) {
											onCopyRecordUrl(ids);
										}
									}}
									onActivate={() => {
										const ids = getSelectedIds();
										if (ids?.length && onActivateRecords) {
											onActivateRecords(ids);
										}
									}}
									onDeactivate={() => {
										const ids = getSelectedIds();
										if (ids?.length && onDeactivateRecords) {
											onDeactivateRecords(ids);
										}
									}}
									onDelete={() => {
										const ids = getSelectedIds();
										// For delete, we allow empty selection to trigger "delete all" dialog
										if (onDeleteRecords) {
											onDeleteRecords(ids || []);
										}
									}}
									onBulkDelete={() => {
										const ids = getSelectedIds();
										if (onBulkDeleteRecords) {
											onBulkDeleteRecords(ids || []);
										}
									}}
									canDelete={canDelete}
									canBulkDelete={canBulkDelete}
									deleteDisabled={isAggregateQuery || hasOneToManyRelationship}
									onRunSpecificWorkflow={(workflow: WorkflowInfo) => {
										const ids = getSelectedIds();
										if (ids?.length && onRunSpecificWorkflow) {
											onRunSpecificWorkflow(workflow, ids);
										}
									}}
									canRunWorkflow={canRunWorkflow}
									workflowDisabled={isAggregateQuery || hasOneToManyRelationship}
									recordActionsDisabledReason={recordActionsDisabledReason}
									onFetchWorkflows={onFetchWorkflows}
									onExport={onExport}
									onExportLocal={onExportLocal}
									canExport={canExport}
									isExporting={isExporting}
									exportDisabledReason={exportDisabledReason}
									entityName={result?.entityLogicalName}
									entityDisplayName={entityDisplayName}
									columns={columnConfig?.columns}
									onReorderColumns={onReorderColumns}
									onRemoveColumns={onRemoveColumns}
									availableAttributes={
										// Root entity attributes to offer, per "Advanced Find only" (#57)
										attributeMetadata && fetchQuery?.entity?.name
											? Array.from(attributeMetadata.get(fetchQuery.entity.name)?.values() || []).filter((a) =>
													isAttributeInScope(a, advancedFindOnly),
												)
											: undefined
									}
									selectedAttributes={fetchQuery?.entity.attributes.map((a) => a.name)}
									onAddColumn={onAddColumn}
									onAddRelatedColumns={onAddRelatedColumns}
									lookupRelationships={lookupRelationships}
									oneToManyRelationships={oneToManyRelationships}
									isLoadingRelationships={isLoadingRelationships}
									onLoadRelatedAttributes={onLoadRelatedAttributes}
									onResetToDefault={onResetToDefault}
									addColumnsDisabledReason={
										isAggregateQuery
											? "Aggregate queries only return Group by and aggregate attributes, with aliases: add them in the tree."
											: undefined
									}
									existingRelatedColumns={existingRelatedColumns}
									isRetrieveAllInProgress={isRetrieveAllInProgress}
									onStopRetrieveAll={onStopRetrieveAll}
								/>
							</div>
							{pagingNotice && (
								<div className={styles.messageBarContainer}>
									<MessageBar intent={pagingNotice.kind === "error" ? "error" : pagingNotice.kind === "stopped" ? "warning" : "info"} layout="multiline">
										<MessageBarBody>
											<MessageBarTitle>
												{pagingNotice.kind === "error"
													? "Couldn't load all records"
													: pagingNotice.kind === "stopped"
														? "Retrieve All stopped"
														: "Paging without a paging cookie"}
											</MessageBarTitle>
											{pagingNotice.kind === "error"
												? `Loaded ${result?.rows.length ?? 0} rows; ${pagingNotice.page ? `page ${pagingNotice.page}` : "the next page"} failed: ${pagingNotice.message}`
												: pagingNotice.kind === "stopped"
													? `Stopped after ${result?.rows.length ?? 0} rows. Scroll to the end of the grid to load more, or execute again.`
													: pagingNotice.message}
										</MessageBarBody>
										<MessageBarActions
											containerAction={
												<Button
													appearance="transparent"
													icon={<Dismiss16Regular />}
													onClick={onDismissPagingNotice}
													aria-label="Dismiss"
												/>
											}
										>
											{pagingNotice.kind === "error" && onRetryPaging && <Button onClick={onRetryPaging}>Retry</Button>}
										</MessageBarActions>
									</MessageBar>
								</div>
							)}
							<div className={styles.gridCard}>
								<ResultsGrid
									result={result}
									isLoading={isExecuting}
									isLoadingMore={isLoadingMore}
									attributeMetadata={attributeMetadata}
									fetchQuery={fetchQuery}
									primaryIdAttribute={primaryIdAttribute}
									onSelectedCountChange={setToolbarSelectedCount}
									onSelectionChange={handleSelectionChange}
									columnConfig={columnConfig}
									onColumnResize={onColumnResize}
									onSortChange={onSortChange}
									onLoadMore={onLoadMore}
									displaySettings={displaySettings}
								/>
							</div>
						</div>
					)}
				</div>
			</div>
		</EditorSyncContext.Provider>
	);
}
