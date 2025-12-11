/**
 * Tabbed preview panel with FetchXML editor and Results grid
 */

import { useState, type ReactNode } from "react";
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
import { Play24Regular, Dismiss16Regular } from "@fluentui/react-icons";
import { FetchXmlEditor } from "./FetchXmlEditor";
import { ResultsGrid, type QueryResult, type SortChangeData } from "./ResultsGrid";
import { ResultsCommandBar } from "./ResultsCommandBar";
import type { AttributeMetadata } from "../../api/pptbClient";
import type { FetchNode } from "../../model/nodes";
import type { ParseResult } from "../../model/fetchxmlParser";
import type { LayoutXmlConfig } from "../../model/layoutxml";

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
	},
	// Grid card styling
	gridCard: {
		borderRadius: tokens.borderRadiusMedium,
		boxShadow: tokens.shadow16,
		backgroundColor: tokens.colorNeutralBackground1,
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		display: "flex",
		flexDirection: "column",
		minHeight: "240px",
		minWidth: "480px",
		overflow: "hidden",
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
		display: "grid",
		gridTemplateRows: "auto 8px 1fr",
		height: "100%",
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
	result: QueryResult | null;
	isExecuting?: boolean;
	/** Whether more pages are being loaded (for progress display) */
	isLoadingMore?: boolean;
	onExecute?: () => void;
	onExport?: () => void;
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
}

export function PreviewTabs({
	xml,
	result,
	isExecuting,
	isLoadingMore,
	onExecute,
	onExport,
	onParseToTree,
	attributeMetadata,
	fetchQuery,
	columnConfig,
	onColumnResize,
	onReorderColumns,
	onAddColumn,
	onSortChange,
	saveViewButton,
	onLoadMore,
	canExport,
	isExporting,
	exportError,
	onDismissExportError,
	exportDisabledReason,
}: PreviewTabsProps) {
	const styles = useStyles();
	const [selectedTab, setSelectedTab] = useState<"xml" | "results">("xml");
	const [toolbarSelectedCount, setToolbarSelectedCount] = useState(0);

	const handleTabSelect = (_event: SelectTabEvent, data: SelectTabData) => {
		setSelectedTab(data.value as "xml" | "results");
	};

	const handleExecute = () => {
		// Switch to Results tab when executing
		setSelectedTab("results");
		onExecute?.();
	};

	return (
		<div className={styles.container}>
			<div className={styles.header}>
				<TabList selectedValue={selectedTab} onTabSelect={handleTabSelect}>
					<Tab value="xml">FetchXML</Tab>
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
					<ToolbarDivider />
					{saveViewButton}
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
				{selectedTab === "results" && (
					<div className={styles.resultsLayout}>
						<div className={styles.toolbarCard}>
							<ResultsCommandBar
								selectedCount={toolbarSelectedCount}
								onOpen={() => console.log("Open not yet implemented")}
								onCopyUrl={() => console.log("Copy URL not yet implemented")}
								onActivate={() => console.log("Activate not yet implemented")}
								onDeactivate={() => console.log("Deactivate not yet implemented")}
								onDelete={() => console.log("Delete not yet implemented")}
								onExport={onExport || (() => console.log("Export not yet implemented"))}
								canExport={canExport}
								isExporting={isExporting}
								exportDisabledReason={exportDisabledReason}
								entityName={result?.entityLogicalName}
								columns={columnConfig?.columns}
								onReorderColumns={onReorderColumns}
								availableAttributes={
									// Get root entity attributes from multi-entity map
									attributeMetadata && fetchQuery?.entity?.name
										? Array.from(attributeMetadata.get(fetchQuery.entity.name)?.values() || [])
										: undefined
								}
								selectedAttributes={fetchQuery?.entity.attributes.map((a) => a.name)}
								onAddColumn={onAddColumn}
							/>
						</div>
						<div /> {/* 8px spacer */}
						<div className={styles.gridCard}>
							<ResultsGrid
								result={result}
								isLoading={isExecuting}
								isLoadingMore={isLoadingMore}
								attributeMetadata={attributeMetadata}
								fetchQuery={fetchQuery}
								onSelectedCountChange={setToolbarSelectedCount}
								columnConfig={columnConfig}
								onColumnResize={onColumnResize}
								onSortChange={onSortChange}
								onLoadMore={onLoadMore}
							/>
						</div>
					</div>
				)}
			</div>
		</div>
	);
}
