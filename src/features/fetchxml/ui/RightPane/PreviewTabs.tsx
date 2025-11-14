/**
 * Tabbed preview panel with FetchXML editor and Results grid
 */

import { useState } from "react";
import {
	TabList,
	Tab,
	type SelectTabData,
	type SelectTabEvent,
	Button,
	makeStyles,
	Toolbar,
	ProgressBar,
	tokens,
} from "@fluentui/react-components";
import { Play24Regular, ArrowDownload24Regular } from "@fluentui/react-icons";
import { XmlTextArea } from "./XmlTextArea";
import { ResultsGrid, type QueryResult } from "./ResultsGrid";

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
		padding: "8px",
		backgroundColor: tokens.colorNeutralBackground1,
	},
	contentWrapper: {
		height: "100%",
		overflow: "hidden",
		border: `1px solid ${tokens.colorNeutralStroke2}`,
		borderRadius: "4px",
	},
});

interface PreviewTabsProps {
	xml: string;
	result: QueryResult | null;
	isExecuting?: boolean;
	onExecute?: () => void;
	onExport?: () => void;
}

export function PreviewTabs({ xml, result, isExecuting, onExecute, onExport }: PreviewTabsProps) {
	const styles = useStyles();
	const [selectedTab, setSelectedTab] = useState<"xml" | "results">("xml");

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
						disabled={isExecuting || !xml}
					>
						{isExecuting ? "Executing..." : "Execute"}
					</Button>
					{selectedTab === "results" && result && result.rows.length > 0 && (
						<Button appearance="subtle" icon={<ArrowDownload24Regular />} onClick={onExport}>
							Export
						</Button>
					)}
				</Toolbar>
			</div>
			{isExecuting && (
				<div className={styles.progressContainer}>
					<ProgressBar />
				</div>
			)}
			<div className={styles.tabContent}>
				<div className={styles.contentWrapper}>
					{selectedTab === "xml" && <XmlTextArea xml={xml} />}
					{selectedTab === "results" && <ResultsGrid result={result} isLoading={isExecuting} />}
				</div>
			</div>
		</div>
	);
}
