/**
 * Tabbed preview panel with FetchXML editor and Results grid
 */

import { useState, useCallback, type ReactNode } from "react";
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
	tokens,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	MessageBarActions,
} from "@fluentui/react-components";
import { Play24Regular, Dismiss16Regular, Settings20Regular } from "@fluentui/react-icons";
import { FetchXmlEditor } from "./FetchXmlEditor";
import { LayoutXmlViewer } from "./LayoutXmlViewer";
import { ResultsGrid, type QueryResult, type SortChangeData } from "./ResultsGrid";
import { ResultsCommandBar } from "./ResultsCommandBar";
import type { RelatedEntityColumn } from "./AddColumnsPanel";
import type { AttributeMetadata, RelationshipMetadata } from "../../api/pptbClient";
import type { FetchNode } from "../../model/nodes";
import type { ParseResult } from "../../model/fetchxmlParser";
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

interface PreviewTabsProps {
	xml: string;
	/** LayoutXML string for the grid column configuration */
	layoutXml?: string;
	result: QueryResult | null;
	isExecuting?: boolean;
	/** Whether more pages are being loaded (for progress display) */
	isLoadingMore?: boolean;
	onExecute?: () => void;
	/** Export via Dataverse ExportToExcel API */
	onExport?: () => void;
	/** Export locally using exceljs */
	onExportLocal?: () => void;
	onParseToTree?: (xmlString: string) => ParseResult;
	/** Multi-entity attribute metadata: Map<entityLogicalName, Map<attributeLogicalName, AttributeMetadata>> */
	attributeMetadata?: Map<string, Map<string, AttributeMetadata>>;
	fetchQuery?: FetchNode | null;
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
	onRemoveColumn?: (columnName: string) => void;
	/** Callback when user changes sort on a column */
	onSortChange?: (data: SortChangeData) => void;
	/** Optional SaveViewButton to render in the toolbar */
	saveViewButton?: ReactNode;
	/** Callback when user scrolls near bottom (infinite scroll) */
	onLoadMore?: () => void;
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
	/** Callback when selection changes in ResultsGrid */
	onSelectionChange?: (recordIds: string[]) => void;
	/** Callback to get currently selected record IDs */
	getSelectedRecordIds?: () => string[];
	/** Callback when user clicks Settings button */
	onOpenSettings?: () => void;
	/** Display settings (logical names, value format) */
	displaySettings?: DisplaySettings;
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
	onRemoveColumn,
	onSortChange,
	saveViewButton,
	onLoadMore,
	canExport,
	isExporting,
	exportError,
	onDismissExportError,
	exportDisabledReason,
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
	onSelectionChange,
	getSelectedRecordIds,
	onOpenSettings,
	displaySettings,
}: PreviewTabsProps) {
	const styles = useStyles();
	const [selectedTab, setSelectedTab] = useState<"xml" | "layout" | "results">("xml");
	const [toolbarSelectedCount, setToolbarSelectedCount] = useState(0);
	const [selectedRecordIds, setSelectedRecordIds] = useState<string[]>([]);

	const handleTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
		setSelectedTab(data.value as "xml" | "layout" | "results");
	};

	const handleExecute = () => {
		// Switch to Results tab when executing
		setSelectedTab("results");
		onExecute?.();
	};

	const handleSelectionChange = useCallback(
		(recordIds: string[]) => {
			setSelectedRecordIds(recordIds);
			// Also notify parent
			onSelectionChange?.(recordIds);
		},
		[onSelectionChange]
	);

	// If parent provided a getter, use ours; otherwise use internal state
	const getSelectedIds = useCallback(() => {
		return getSelectedRecordIds ? getSelectedRecordIds() : selectedRecordIds;
	}, [getSelectedRecordIds, selectedRecordIds]);

	return (
		<div className={styles.container}>
			<div className={styles.header}>
				<TabList selectedValue={selectedTab} onTabSelect={handleTabSelect}>
					<Tab value="xml">FetchXML</Tab>
					<Tab value="layout">LayoutXML</Tab>
					<Tab value="results">Results</Tab>
				</TabList>
				<Toolbar size="small">
					<Button
						appearance="primary"
						icon={<Play24Regular />}
						onClick={handleExecute}
						disabled={isExecuting || isLoadingMore || !xml}
					>
						{isExecuting ? "Executing..." : "Execute"}
					</Button>
					{selectedTab === "xml" && saveViewButton && (
						<>
							<ToolbarDivider />
							{saveViewButton}
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
			<div className={styles.tabContent}>
				{selectedTab === "xml" && (
					<div className={styles.codeCard}>
						<FetchXmlEditor xml={xml} onParseToTree={onParseToTree} />
					</div>
				)}
				{selectedTab === "layout" && (
					<div className={styles.codeCard}>
						<LayoutXmlViewer layoutXml={layoutXml || ""} />
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
								onFetchWorkflows={onFetchWorkflows}
								onExport={onExport || (() => console.log("Export not yet implemented"))}
								onExportLocal={onExportLocal}
								canExport={canExport}
								isExporting={isExporting}
								exportDisabledReason={exportDisabledReason}
								entityName={result?.entityLogicalName}
								entityDisplayName={entityDisplayName}
								columns={columnConfig?.columns}
								onReorderColumns={onReorderColumns}
								onRemoveColumn={onRemoveColumn}
								availableAttributes={
									// Get root entity attributes from multi-entity map
									attributeMetadata && fetchQuery?.entity?.name
										? Array.from(attributeMetadata.get(fetchQuery.entity.name)?.values() || [])
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
							/>
						</div>
						<div className={styles.gridCard}>
							<ResultsGrid
								result={result}
								isLoading={isExecuting}
								isLoadingMore={isLoadingMore}
								attributeMetadata={attributeMetadata}
								fetchQuery={fetchQuery}
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
	);
}
