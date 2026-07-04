/**
 * Code Generation Panel — renders 7 code-format tabs from a FetchXML query.
 *
 * Sync formats (C# FetchExpression, JavaScript, pac CLI, Power Automate, Web API)
 * are computed immediately. Async formats (C# QueryExpression, SQL) call Dataverse
 * Web API functions via pptbClient and display a Spinner while loading.
 */

import { useState, useEffect, useCallback, useRef } from "react";
import Editor from "@monaco-editor/react";
import {
	Button,
	TabList,
	Tab,
	type SelectTabData,
	type SelectTabEvent,
	Spinner,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Badge,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Copy20Regular, Info20Regular } from "@fluentui/react-icons";
import { useTheme } from "../../../../shared/contexts/ThemeContext";
import { usePptbContext } from "../../../../shared/hooks/usePptbContext";
import {
	generateFxsSyncCode,
	generatePowerAutomateSpec,
	FXS_SYNC_FORMAT_LABELS,
	FXS_SYNC_FORMAT_LANG,
	type FxsSyncFormat,
} from "../../engine/fetchxmlCodeGenerators";
import { PowerAutomatePane } from "./PowerAutomatePane";
import { generateCSharpQueryExpression } from "../../engine/queryExpressionCodegen";
import type { QeQueryExpression } from "../../engine/queryExpressionTypes";
import { fetchXmlToQueryExpression, fetchXmlToSQL } from "../../api/pptbClient";
import { loadEntityMetadata } from "../../api/dataverseMetadata";

// ─────────────────────────────────────────────────────────────────────────────
// Styles
// ─────────────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
	root: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		overflow: "hidden",
	},
	header: {
		display: "flex",
		alignItems: "center",
		justifyContent: "space-between",
		borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
		paddingLeft: "8px",
		paddingRight: "8px",
		paddingTop: "4px",
		paddingBottom: "4px",
		flexShrink: 0,
		backgroundColor: tokens.colorNeutralBackground2,
	},
	tabListWrap: {
		display: "flex",
		alignItems: "center",
		gap: "4px",
		flexWrap: "wrap",
	},
	tabLabel: {
		display: "flex",
		alignItems: "center",
		gap: "4px",
	},
	copyButton: {
		flexShrink: 0,
	},
	body: {
		flex: 1,
		overflow: "hidden",
		position: "relative",
	},
	paPane: {
		height: "100%",
		overflowY: "auto",
		padding: "16px",
		boxSizing: "border-box",
	},
	centerBox: {
		display: "flex",
		flexDirection: "column",
		alignItems: "center",
		justifyContent: "center",
		height: "100%",
		gap: "12px",
		padding: "24px",
	},
	offlineText: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
		textAlign: "center",
		maxWidth: "320px",
	},
	messagePad: {
		padding: "12px",
	},
});

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

type CodeFormat = FxsSyncFormat | "csharp-queryexpr" | "sql";

const ALL_FORMATS: CodeFormat[] = [
	"csharp-fetchexpr",
	"csharp-queryexpr",
	"sql",
	"javascript",
	"pac",
	"powerautomate",
	"webapi",
];

const FORMAT_LABELS: Record<CodeFormat, string> = {
	...FXS_SYNC_FORMAT_LABELS,
	"csharp-queryexpr": "C# QueryExpression",
	sql: "SQL",
};

const FORMAT_LANG: Record<CodeFormat, string> = {
	...FXS_SYNC_FORMAT_LANG,
	"csharp-queryexpr": "csharp",
	sql: "sql",
};

// ─────────────────────────────────────────────────────────────────────────────
// Component
// ─────────────────────────────────────────────────────────────────────────────

interface CodePanelProps {
	fetchXml: string;
}

export function CodePanel({ fetchXml }: CodePanelProps) {
	const styles = useStyles();
	const { isDark } = useTheme();
	const { connected, environmentUrl } = usePptbContext();

	const [selectedFormat, setSelectedFormat] = useState<CodeFormat>("csharp-fetchexpr");
	const [copied, setCopied] = useState(false);

	// Async: C# QueryExpression
	const [qeCode, setQeCode] = useState<string>("");
	const [qeLoading, setQeLoading] = useState(false);
	const [qeError, setQeError] = useState<string | null>(null);

	// Async: SQL
	const [sqlCode, setSqlCode] = useState<string>("");
	const [sqlLoading, setSqlLoading] = useState(false);
	const [sqlError, setSqlError] = useState<string | null>(null);

	// Power Automate: resolved entity set name from metadata
	const [paEntitySetName, setPaEntitySetName] = useState<string | undefined>(undefined);

	// Abort / cancel async calls when fetchXml changes
	const abortRef = useRef<{ cancelled: boolean }>({ cancelled: false });

	useEffect(() => {
		// Cancel any in-flight previous calls
		abortRef.current.cancelled = true;
		const signal = { cancelled: false };
		abortRef.current = signal;

		if (!fetchXml || !connected) {
			setQeCode("");
			setQeError(null);
			setSqlCode("");
			setSqlError(null);
			setPaEntitySetName(undefined);
			return;
		}

		// Resolve the real OData entity set name for Power Automate
		const entityNameMatch = fetchXml.match(/entity\s+name=["']([^"']+)["']/i);
		if (entityNameMatch) {
			loadEntityMetadata(entityNameMatch[1])
				.then((meta) => {
					if (!signal.cancelled) setPaEntitySetName(meta.EntitySetName);
				})
				.catch(() => {
					/* silently fall back to heuristic */
				});
		}

		// Trigger both async generators in parallel
		setQeLoading(true);
		setQeError(null);
		setSqlLoading(true);
		setSqlError(null);

		// C# QueryExpression
		fetchXmlToQueryExpression(fetchXml)
			.then((qe) => {
				if (signal.cancelled) return;
				const code = generateCSharpQueryExpression(qe as unknown as QeQueryExpression);
				setQeCode(code);
			})
			.catch((err: unknown) => {
				if (signal.cancelled) return;
				setQeError(err instanceof Error ? err.message : String(err));
				setQeCode("");
			})
			.finally(() => {
				if (!signal.cancelled) setQeLoading(false);
			});

		// SQL
		fetchXmlToSQL(fetchXml)
			.then((sql) => {
				if (signal.cancelled) return;
				setSqlCode(sql);
			})
			.catch((err: unknown) => {
				if (signal.cancelled) return;
				setSqlError(err instanceof Error ? err.message : String(err));
				setSqlCode("");
			})
			.finally(() => {
				if (!signal.cancelled) setSqlLoading(false);
			});
	}, [fetchXml, connected]);

	const handleTabSelect = (_e: SelectTabEvent, data: SelectTabData) => {
		setSelectedFormat(data.value as CodeFormat);
	};

	// Resolve the code to display for the current tab
	const currentCode = useCallback((): string => {
		switch (selectedFormat) {
			case "csharp-queryexpr":
				return qeCode;
			case "sql":
				return sqlCode;
			default:
				return generateFxsSyncCode(
					selectedFormat as FxsSyncFormat,
					fetchXml,
					selectedFormat === "webapi" ? paEntitySetName : undefined,
					environmentUrl,
				);
		}
	}, [selectedFormat, fetchXml, qeCode, sqlCode, environmentUrl]);

	const handleCopy = useCallback(() => {
		const code = currentCode();
		if (!code) return;
		void navigator.clipboard.writeText(code).then(() => {
			setCopied(true);
			setTimeout(() => setCopied(false), 1500);
		});
	}, [currentCode]);

	const isAsyncLoading =
		(selectedFormat === "csharp-queryexpr" && qeLoading) ||
		(selectedFormat === "sql" && sqlLoading);

	const asyncError =
		selectedFormat === "csharp-queryexpr" ? qeError : selectedFormat === "sql" ? sqlError : null;

	const showOffline =
		(selectedFormat === "csharp-queryexpr" || selectedFormat === "sql") && !connected;

	const code = currentCode();
	const monacoTheme = isDark ? "vs-dark" : "light";
	const monacoLang = FORMAT_LANG[selectedFormat];

	return (
		<div className={styles.root}>
			{/* Header: tab bar + copy button */}
			<div className={styles.header}>
				<div className={styles.tabListWrap}>
					<TabList
						size="small"
						appearance="subtle"
						selectedValue={selectedFormat}
						onTabSelect={handleTabSelect}
					>
						{ALL_FORMATS.map((fmt) => (
							<Tab key={fmt} value={fmt}>
								<span className={styles.tabLabel}>
									{FORMAT_LABELS[fmt]}
									{fmt === "sql" && (
										<Tooltip
											content="Undocumented API — may change without notice"
											relationship="label"
										>
											<Badge
												size="small"
												appearance="tint"
												color="warning"
												style={{ cursor: "default" }}
											>
												Preview
											</Badge>
										</Tooltip>
									)}
									{(fmt === "csharp-queryexpr" || fmt === "sql") && !connected && (
										<Tooltip content="Connect to an environment to generate" relationship="label">
											<Info20Regular style={{ opacity: 0.5, width: 14, height: 14 }} />
										</Tooltip>
									)}
								</span>
							</Tab>
						))}
					</TabList>
				</div>
				<Tooltip
					content={
						selectedFormat === "powerautomate"
							? "Use the per-field Copy buttons below"
							: copied
								? "Copied!"
								: "Copy to clipboard"
					}
					relationship="label"
				>
					<Button
						appearance="subtle"
						size="small"
						icon={<Copy20Regular />}
						className={styles.copyButton}
						onClick={handleCopy}
						disabled={isAsyncLoading || showOffline || !code || selectedFormat === "powerautomate"}
						aria-label="Copy code to clipboard"
					>
						{copied ? "Copied!" : "Copy"}
					</Button>
				</Tooltip>
			</div>

			{/* Body */}
			<div className={styles.body}>
				{/* Spinner for async tabs */}
				{isAsyncLoading && (
					<div className={styles.centerBox}>
						<Spinner size="medium" label="Generating..." />
					</div>
				)}

				{/* Error state for async tabs */}
				{!isAsyncLoading && asyncError && (
					<div className={styles.messagePad}>
						<MessageBar intent="error">
							<MessageBarBody>
								<MessageBarTitle>Generation failed</MessageBarTitle>
								{asyncError}
							</MessageBarBody>
						</MessageBar>
					</div>
				)}

				{/* Offline state for async tabs */}
				{!isAsyncLoading && showOffline && (
					<div className={styles.centerBox}>
						<Info20Regular style={{ opacity: 0.4, width: 32, height: 32 }} />
						<p className={styles.offlineText}>
							Connect to a Dataverse environment to generate{" "}
							{selectedFormat === "csharp-queryexpr" ? "C# QueryExpression" : "SQL"} code.
						</p>
					</div>
				)}

				{/* Power Automate pane — field-by-field form */}
				{!isAsyncLoading && !asyncError && !showOffline && selectedFormat === "powerautomate" && (
					<div className={styles.paPane}>
						<PowerAutomatePane spec={generatePowerAutomateSpec(fetchXml, paEntitySetName)} />
					</div>
				)}

				{/* Monaco editor for all other formats */}
				{!isAsyncLoading && !asyncError && !showOffline && selectedFormat !== "powerautomate" && (
					<Editor
						height="100%"
						language={monacoLang}
						value={code}
						theme={monacoTheme}
						options={{
							readOnly: true,
							minimap: { enabled: false },
							scrollBeyondLastLine: false,
							wordWrap: selectedFormat === "webapi" ? "on" : "off",
							fontSize: 13,
							lineNumbers: "on",
							renderLineHighlight: "none",
							folding: true,
							automaticLayout: true,
						}}
					/>
				)}
			</div>
		</div>
	);
}
