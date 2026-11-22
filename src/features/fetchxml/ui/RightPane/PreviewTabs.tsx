/**
 * Tabbed preview panel with FetchXML editor and Results grid
 */

import { useState, useCallback, useMemo, useRef, type ReactNode } from "react";
import { makeStyles, ProgressBar, tokens } from "@fluentui/react-components";
import { FetchXmlEditor } from "./FetchXmlEditor";
import { CodePanel } from "./CodePanel";
import { LayoutXmlViewer } from "./LayoutXmlViewer";
import {
	ResultsGrid,
	type QueryResult,
	type ResultsGridHandle,
	type SortChangeData,
} from "./ResultsGrid";
import { ResultsCommandBar, type CloseMenuItem } from "./ResultsCommandBar";
import type { RelatedEntityColumn } from "./AddColumnsPanel";
import type { PagingNotice } from "../../api/paging";
import { isAttributeInScope, useAdvancedFindOnly } from "../../state/metadataScope";
import { NON_RETURNING_LINK_TYPES } from "../../model/fetchxmlRules";
import type { AttributeMetadata, RelationshipMetadata } from "../../api/pptbClient";
import type { ConditionNode, FetchNode } from "../../model/nodes";
import type { NewCondition } from "../../model/columnFilters";
import type { ParseResult } from "../../model/fetchxmlParser";
import { parseFetchXml } from "../../model/fetchxmlParser";
import { generateFetchXml } from "../../model/fetchxml";
import { DATAVERSE_DEFAULT_PAGE_SIZE } from "../../api/pageRecovery";
import type { LoadFetchXmlOptions } from "../../state/builderContext";
import { ParseChangesDialog } from "./ParseChangesDialog";
import { PaneErrorBoundary } from "../../../../shared/components/ErrorBoundary";
import { ResultsTabBars, type ExportError } from "./ResultsTabBars";
import type { LayoutXmlConfig } from "../../model/layoutxml";
import { isDataverseAvailable, type WorkflowInfo } from "../../api/pptbClient";
import {
	normalizeGridDensity,
	type DisplaySettings,
	type GridDensity,
} from "../../model/displaySettings";
import type { NodeId } from "../../model/nodes";
import { collectAdvisories, mergeDataverseCheck, type Advisory } from "../../model/advisories";
import { AdvisoriesPanel } from "./AdvisoriesPanel";
import { useAdvisorMetadata, useDataverseCheck } from "./useAdvisories";

import { PreviewHeader } from "./PreviewHeader";
import type { EditorSession } from "../../../../app/hooks/useEditorSession";

/** FetchXML as the tree writes it, to compare meaning rather than formatting; null if it doesn't parse */
function parseForAdvisories(xmlText: string): { fetchNode: FetchNode; canonical: string } | null {
	const result = parseFetchXml(xmlText);
	return result.success && result.fetchNode
		? { fetchNode: result.fetchNode, canonical: generateFetchXml(result.fetchNode) }
		: null;
}

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		overflow: "hidden",
	},
	progressContainer: {
		paddingLeft: "16px",
		paddingRight: "16px",
		paddingTop: "8px",
		paddingBottom: "8px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
	},
	// Each tab's content sits flat on the pane (#233): no canvas, no floating cards
	tabContent: {
		flex: 1,
		overflow: "hidden",
		backgroundColor: tokens.colorNeutralBackground1,
	},
	// The Results command bar, divided from the grid by a thin line
	toolbarCard: {
		padding: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalS}`,
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
		flexShrink: 0,
	},
	// The results grid takes the rest of the tab
	gridCard: {
		display: "flex",
		flexDirection: "column",
		flex: 1,
		minHeight: "240px",
		minWidth: 0,
		overflow: "hidden",
	},
	// A full-height tab body (FetchXML, LayoutXML, Code)
	codeCard: {
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

interface PreviewTabsProps {
	/** The editor session and Run, shared with the app bar (#278) */
	session: EditorSession;
	/** The FetchXML editor's Parse to tree */
	onParseToTree?: (xmlString: string, options?: LoadFetchXmlOptions) => ParseResult;
	xml: string;
	/** LayoutXML string for the grid column configuration */
	layoutXml?: string;
	result: QueryResult | null;
	isExecuting?: boolean;
	/** Whether more pages are being loaded (for progress display) */
	isLoadingMore?: boolean;
	/** Export via Dataverse ExportToExcel API */
	onExport?: () => void;
	/** Export locally using exceljs */
	onExportLocal?: () => void;
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
	/** The results grid's column filters (#23) */
	columnFilters?: ConditionNode[];
	onAddColumnFilter?: (condition: NewCondition) => void;
	onRemoveColumnFilters?: (conditionIds: string[]) => void;
	/** Quick find test value field, shown under the toolbar while the query has {0} (#142) */
	quickFindTestValue?: ReactNode;
	/** Callback when user scrolls near bottom (infinite scroll) */
	onLoadMore?: () => void;
	/** Paging notice above the grid: failed page, Stop, simple paging (#60) */
	pagingNotice?: PagingNotice | null;
	/** Retrieve All is waiting this long to retry a throttled page (#183) */
	pagingWaitSeconds?: number;
	isRetrieveAllInProgress?: boolean;
	onStopRetrieveAll?: () => void;
	onRetryPaging?: () => void;
	onDismissPagingNotice?: () => void;
	/** Whether export is available (requires a saved view) */
	canExport?: boolean;
	/** Whether export is in progress */
	isExporting?: boolean;
	/** Why the last export failed, and whether Try again can repeat it (#277) */
	exportError?: ExportError;
	/** Callback to dismiss export error */
	onDismissExportError?: () => void;
	/** Repeats the export that failed (#277) */
	onRetryExport?: () => void;
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
	/** Why Activate is off for this entity (#143) */
	activateDisabledReason?: string;
	/** Why Deactivate is off for this entity (#143) */
	deactivateDisabledReason?: string;
	/** Close menu items for close-message entities (#149); replaces Deactivate */
	closeActions?: CloseMenuItem[];
	/** Why the Close menu is off (#149) */
	closeDisabledReason?: string;
	/** Callback when user picks a close action, with the selected record IDs (#149) */
	onCloseRecords?: (actionKey: string, recordIds: string[]) => void;
	/** Callback when user wants to run a specific workflow directly */
	onRunSpecificWorkflow?: (workflow: WorkflowInfo, recordIds: string[]) => void;
	/** Whether user can delete records */
	canDelete?: boolean;
	/** Whether user can bulk delete records */
	canBulkDelete?: boolean;
	/** Whether user can run workflows */
	canRunWorkflow?: boolean;
	/** Fetch available workflows for the entity */
	onFetchWorkflows?: (entityLogicalName?: string) => Promise<WorkflowInfo[]>;
	/** Whether the current query is an aggregate query (disables delete/workflow) */
	isAggregateQuery?: boolean;
	/** Why Delete is off for this query's shape (aggregate, or rows that repeat) (#232) */
	deleteDisabledReason?: string;
	/** Why Run workflow is off for this query's shape (#232) */
	workflowDisabledReason?: string;
	/** Activate / Deactivate / Close show: the Write check finished and didn't say no (#232) */
	canChangeState?: boolean;
	/** When set, results have no record IDs and all record actions are disabled with this explanation */
	recordActionsDisabledReason?: string;
	/** Why "Delete all" (no selection) is off: the query changed since the results loaded, or nothing ran (#174) */
	deleteAllDisabledReason?: string;
	/** Callback when selection changes in ResultsGrid */
	onSelectionChange?: (recordIds: string[]) => void;
	/** Callback to get currently selected record IDs */
	getSelectedRecordIds?: () => string[];
	/** Display settings (logical names, value format) */
	displaySettings?: DisplaySettings;
	/** Change the grid and tree density (#228) */
	onGridDensityChange?: (density: GridDensity) => void;
	/** Primary key attribute from Dataverse entity metadata (e.g. "activityid" for activity entities) */
	primaryIdAttribute?: string;
	/** Identity of the tool's connection: "" with none (usePptbContext); Check with Dataverse needs one (#97) */
	connectionKey?: string;
	/** The tree's selected node, highlighted in the Advisories tab (#97) */
	selectedNodeId?: NodeId | null;
	/** Select a node in the tree: an advisory's location (#97) */
	onSelectNode?: (nodeId: NodeId) => void;
	/** Add an order on an attribute: the Advisories tab's quick fix (#97) */
	onAddOrder?: (parentId: NodeId, attribute: string) => void;
}

export function PreviewTabs({
	xml,
	layoutXml,
	result,
	isExecuting,
	isLoadingMore,
	onExport,
	onExportLocal,
	attributeMetadata,
	fetchQuery,
	columnConfig,
	onColumnResize,
	onReorderColumns,
	onAddColumn,
	onAddRelatedColumns,
	onRemoveColumns,
	onSortChange,
	columnFilters,
	onAddColumnFilter,
	onRemoveColumnFilters,
	quickFindTestValue,
	onLoadMore,
	pagingNotice,
	pagingWaitSeconds,
	isRetrieveAllInProgress,
	onStopRetrieveAll,
	onRetryPaging,
	onDismissPagingNotice,
	canExport,
	isExporting,
	exportError,
	onDismissExportError,
	onRetryExport,
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
	activateDisabledReason,
	deactivateDisabledReason,
	closeActions,
	closeDisabledReason,
	onCloseRecords,
	onRunSpecificWorkflow,
	canDelete,
	canBulkDelete,
	canRunWorkflow,
	onFetchWorkflows,
	isAggregateQuery,
	deleteDisabledReason,
	workflowDisabledReason,
	canChangeState,
	recordActionsDisabledReason,
	deleteAllDisabledReason,
	onSelectionChange,
	getSelectedRecordIds,
	displaySettings,
	onGridDensityChange,
	primaryIdAttribute,
	connectionKey,
	selectedNodeId,
	onSelectNode,
	onAddOrder,
	session,
	onParseToTree,
}: PreviewTabsProps) {
	const styles = useStyles();
	const advancedFindOnly = useAdvancedFindOnly();
	// The editor and Run live in AppShell's editor session, shared with the app bar (#278)
	const {
		selectedTab,
		setSelectedTab,
		isEditorModeActive,
		editorText,
		flushEditorText,
		fetchXmlEditorRef,
		editorReplace,
		pendingSync,
		resolvePendingSync,
		onEditorStateChange,
		run,
	} = session;
	const [toolbarSelectedCount, setToolbarSelectedCount] = useState(0);
	const [selectedRecordIds, setSelectedRecordIds] = useState<string[]>([]);

	// The command bar's Copy menu copies through the grid (#50)
	const resultsGridRef = useRef<ResultsGridHandle>(null);
	// Quick search over the loaded rows (#23), and how many rows it shows
	const [searchText, setSearchText] = useState("");
	const [visibleRowCount, setVisibleRowCount] = useState(0);
	// Related columns already returned by root link-entities Add columns would reuse (#114)
	const existingRelatedColumns = useMemo(
		() =>
			new Set(
				(fetchQuery?.entity.links ?? [])
					.filter((link) => !NON_RETURNING_LINK_TYPES.has(link.linkType))
					.flatMap((link) =>
						link.attributes.map(
							(a) => `${link.name}|${link.from}|${link.to}|${a.name}`,
						),
					),
			),
		[fetchQuery],
	);

	// ---- Advisories (#97): what the user is looking at, the editor's FetchXML in editor mode
	const advisorySubject = useMemo(() => {
		if (!isEditorModeActive)
			return { fetch: fetchQuery ?? null, key: xml, inTree: true, editorInvalid: false };
		const parsed = parseForAdvisories(editorText);
		if (!parsed)
			return {
				fetch: null,
				key: editorText,
				inTree: false,
				editorInvalid: editorText.trim() !== "",
			};
		// The editor means the same as the tree: use the tree's nodes, so locations and fixes work
		if (fetchQuery && parsed.canonical === xml)
			return { fetch: fetchQuery, key: xml, inTree: true, editorInvalid: false };
		return {
			fetch: parsed.fetchNode,
			key: parsed.canonical,
			inTree: false,
			editorInvalid: false,
		};
	}, [isEditorModeActive, editorText, fetchQuery, xml]);
	// The metadata rules' metadata loads from the tab's first open (#170)
	const [advisoriesOpened, setAdvisoriesOpened] = useState(false);
	if (selectedTab === "advisories" && !advisoriesOpened) setAdvisoriesOpened(true);
	const advisorMetadata = useAdvisorMetadata(advisorySubject.fetch, advisoriesOpened);
	const dataverseCheck = useDataverseCheck();
	const checkResult = dataverseCheck.result;
	const checkStale = !!checkResult && checkResult.key !== advisorySubject.key;
	// The results of this query are paged: more pages loaded or available (Load more, Retrieve All, #170)
	const resultQueryXml = useMemo(
		() => (result?.query ? generateFetchXml(result.query) : undefined),
		[result?.query],
	);
	const resultsPaged =
		!!result &&
		(result.moreRecords === true || result.rows.length > DATAVERSE_DEFAULT_PAGE_SIZE) &&
		resultQueryXml === advisorySubject.key;
	const { advisories, added } = useMemo((): { advisories: Advisory[]; added: number } => {
		const subject = advisorySubject.fetch;
		if (!subject) return { advisories: [], added: 0 };
		const own = collectAdvisories(subject, advisorMetadata, { resultsPaged });
		// A check of an earlier query isn't shown or counted until Check again (#170); its status
		// line still says when it was
		return checkResult && !checkStale
			? mergeDataverseCheck(subject, own, checkResult.check)
			: { advisories: own, added: 0 };
	}, [advisorySubject.fetch, advisorMetadata, checkResult, checkStale, resultsPaged]);
	const checkDisabledReason = !isDataverseAvailable()
		? "Dataverse isn't available here. Open FetchXML Studio in Power Platform ToolBox."
		: connectionKey === ""
			? "Connect to an environment to check with Dataverse."
			: undefined;
	/** Check the FetchXML the user is looking at: the editor's text in editor mode, otherwise the tree's */
	const handleCheckWithDataverse = () => {
		if (isEditorModeActive) {
			// Flush the debounce, so the list and the stale check use the text that's checked
			const text = flushEditorText();
			void dataverseCheck.run(text, parseForAdvisories(text)?.canonical ?? text);
		} else {
			void dataverseCheck.run(xml, xml);
		}
	};
	const handleAdvisoryFix = (advisory: Advisory) => {
		if (advisory.fix?.kind === "add-root-order" && fetchQuery && advisorySubject.inTree) {
			onAddOrder?.(fetchQuery.entity.id, advisory.fix.attribute);
		}
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
		<>
			<div className={styles.container}>
				<ParseChangesDialog
					action={pendingSync?.action ?? null}
					changes={pendingSync?.changes ?? []}
					onConfirm={() => resolvePendingSync(true)}
					onCancel={() => resolvePendingSync(false)}
				/>
				<PreviewHeader
					selectedTab={selectedTab}
					onTabSelect={setSelectedTab}
					advisories={advisories}
					rowCount={result?.rows.length}
				/>
				{(isExecuting || isLoadingMore) && (
					<div className={styles.progressContainer}>
						<ProgressBar />
					</div>
				)}
				{quickFindTestValue && (
					<div className={styles.messageBarContainer}>{quickFindTestValue}</div>
				)}
				<div className={styles.tabContent}>
					{/* FetchXML tab — always mounted so Monaco editor state (editor mode, content,
					    cursor position) survives tab switches. Hidden via display:none when inactive;
					    Monaco's automaticLayout re-layouts automatically when it becomes visible. */}
					<div
						className={styles.codeCard}
						style={selectedTab !== "xml" ? { display: "none" } : undefined}
					>
						<PaneErrorBoundary
							where="FetchXML editor"
							title="The FetchXML editor couldn't be shown"
						>
							<FetchXmlEditor
								ref={fetchXmlEditorRef}
								xml={xml}
								onParseToTree={onParseToTree}
								onEditorStateChange={onEditorStateChange}
								replaceRequest={editorReplace}
								resetRequest={session.editorResetId}
								onViewAdvisories={() => setSelectedTab("advisories")}
							/>
						</PaneErrorBoundary>
					</div>
					{selectedTab === "advisories" && (
						<PaneErrorBoundary
							where="Advisories panel"
							title="The advisories couldn't be shown"
							resetKey={advisorySubject.key}
						>
							<AdvisoriesPanel
								advisories={advisories}
								noQuery={!fetchQuery}
								editorInvalid={advisorySubject.editorInvalid}
								inTree={advisorySubject.inTree}
								selectedNodeId={selectedNodeId}
								onSelectNode={onSelectNode}
								onFix={handleAdvisoryFix}
								checking={dataverseCheck.checking}
								checkedAt={checkResult?.checkedAt}
								stale={checkStale}
								added={added}
								checkError={dataverseCheck.error}
								onDismissCheckError={dataverseCheck.dismissError}
								onCheck={handleCheckWithDataverse}
								checkDisabledReason={checkDisabledReason}
							/>
						</PaneErrorBoundary>
					)}
					{selectedTab === "layout" && (
						<div className={styles.codeCard}>
							<PaneErrorBoundary
								where="LayoutXML panel"
								title="The LayoutXML couldn't be shown"
								resetKey={layoutXml}
							>
								<LayoutXmlViewer layoutXml={layoutXml || ""} />
							</PaneErrorBoundary>
						</div>
					)}
					{selectedTab === "code" && (
						<div className={styles.codeCard}>
							<PaneErrorBoundary
								where="Code panel"
								title="The code couldn't be generated"
								resetKey={xml}
							>
								<CodePanel fetchXml={xml} />
							</PaneErrorBoundary>
						</div>
					)}
					{selectedTab === "results" && (
						<PaneErrorBoundary
							where="Results panel"
							title="The results couldn't be shown"
							resetKey={result}
						>
							<div className={styles.resultsLayout}>
								<div className={styles.toolbarCard}>
									<ResultsCommandBar
										selectedCount={toolbarSelectedCount}
										density={normalizeGridDensity(displaySettings?.gridDensity)}
										onDensityChange={onGridDensityChange}
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
										onActivate={
											onActivateRecords &&
											(() => {
												const ids = getSelectedIds();
												if (ids?.length && onActivateRecords) {
													onActivateRecords(ids);
												}
											})
										}
										onDeactivate={
											onDeactivateRecords &&
											(() => {
												const ids = getSelectedIds();
												if (ids?.length && onDeactivateRecords) {
													onDeactivateRecords(ids);
												}
											})
										}
										activateDisabledReason={activateDisabledReason}
										deactivateDisabledReason={deactivateDisabledReason}
										closeActions={closeActions}
										closeDisabledReason={closeDisabledReason}
										onCloseAction={
											onCloseRecords &&
											((key) => {
												const ids = getSelectedIds();
												if (ids?.length) onCloseRecords(key, ids);
											})
										}
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
										deleteDisabledReason={deleteDisabledReason}
										canChangeState={canChangeState}
										onRunSpecificWorkflow={(workflow: WorkflowInfo) => {
											const ids = getSelectedIds();
											if (ids?.length && onRunSpecificWorkflow) {
												onRunSpecificWorkflow(workflow, ids);
											}
										}}
										canRunWorkflow={canRunWorkflow}
										workflowDisabledReason={workflowDisabledReason}
										recordActionsDisabledReason={recordActionsDisabledReason}
										deleteAllDisabledReason={deleteAllDisabledReason}
										onFetchWorkflows={onFetchWorkflows}
										onExport={onExport}
										onExportLocal={onExportLocal}
										loadedCount={
											searchText.trim()
												? visibleRowCount
												: (result?.rows.length ?? 0)
										}
										searchText={searchText}
										onCopyRows={(scope, mode) =>
											resultsGridRef.current?.copyRows(scope, mode)
										}
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
												? Array.from(
														attributeMetadata
															.get(fetchQuery.entity.name)
															?.values() || [],
													).filter((a) =>
														isAttributeInScope(a, advancedFindOnly),
													)
												: undefined
										}
										selectedAttributes={fetchQuery?.entity.attributes.map(
											(a) => a.name,
										)}
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
								<ResultsTabBars
									executeError={executeError}
									onRetryRun={run}
									onDismissExecuteError={onDismissExecuteError}
									isExporting={isExporting}
									exportError={exportError}
									onRetryExport={onRetryExport}
									onDismissExportError={onDismissExportError}
									pagingNotice={pagingNotice}
									rowCount={result?.rows.length ?? 0}
									onRetryPaging={onRetryPaging}
									onDismissPagingNotice={onDismissPagingNotice}
								/>
								<div className={styles.gridCard}>
									<ResultsGrid
										ref={resultsGridRef}
										result={result}
										isLoading={isExecuting}
										isLoadingMore={isLoadingMore}
										throttleWaitSeconds={pagingWaitSeconds}
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
										searchText={searchText}
										onSearchTextChange={setSearchText}
										columnFilters={columnFilters}
										onAddColumnFilter={onAddColumnFilter}
										onRemoveColumnFilters={onRemoveColumnFilters}
										onVisibleCountChange={setVisibleRowCount}
									/>
								</div>
							</div>
						</PaneErrorBoundary>
					)}
				</div>
			</div>
		</>
	);
}
