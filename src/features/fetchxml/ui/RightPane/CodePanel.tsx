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
import { useTheme } from "../../../../shared/contexts/useTheme";
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
import { copyText } from "../../api/host";
import {
	CONVERSION_DELAY_MS,
	conversionKey,
	needsConversion,
	serverConversionFor,
	type ConversionResult,
	type ServerConversion,
} from "./serverConversions";

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

	// Dataverse conversions (C# QueryExpression, SQL): for the selected format only, after typing
	// pauses, and once per query and environment (#199)
	const [conversions, setConversions] = useState<Partial<Record<ServerConversion, ConversionResult>>>({});
	const lastRequestedRef = useRef<Partial<Record<ServerConversion, string>>>({});
	const conversionsRef = useRef(conversions);
	useEffect(() => {
		conversionsRef.current = conversions;
	}, [conversions]);
	const debouncedXml = useDebouncedValue(fetchXml, CONVERSION_DELAY_MS);

	// Power Automate: resolved entity set name from metadata
	const [paEntitySetName, setPaEntitySetName] = useState<string | undefined>(undefined);

	useEffect(() => {
		if (!fetchXml || !connected) {
			setPaEntitySetName(undefined);
			return;
		}
		let cancelled = false;
		// Resolve the real OData entity set name for Power Automate
		const entityNameMatch = fetchXml.match(/<entity\b[^>]*?\sname=["']([^"']+)["']/i);
		if (entityNameMatch) {
			loadEntityMetadata(entityNameMatch[1])
				.then((meta) => {
					if (!cancelled) setPaEntitySetName(meta.EntitySetName);
				})
				.catch(() => {
					/* silently fall back to heuristic */
				});
		}
		return () => {
			cancelled = true;
		};
	}, [fetchXml, connected]);

	useEffect(() => {
		const kind = serverConversionFor(selectedFormat);
		// Wait until typing pauses: until then the debounced query is behind the one on screen
		if (!kind || !debouncedXml || debouncedXml !== fetchXml || !connected) return;
		const key = conversionKey(debouncedXml, environmentUrl);
		if (!needsConversion(key, lastRequestedRef.current[kind], conversionsRef.current[kind])) return;
		lastRequestedRef.current[kind] = key;
		// Retrying a failed conversion: show it loading, not the old error
		if (conversionsRef.current[kind]?.key === key) {
			conversionsRef.current = { ...conversionsRef.current, [kind]: undefined };
			setConversions((prev) => ({ ...prev, [kind]: undefined }));
		}

		const request =
			kind === "queryexpression"
				? fetchXmlToQueryExpression(debouncedXml).then((qe) =>
						generateCSharpQueryExpression(qe as unknown as QeQueryExpression),
					)
				: fetchXmlToSQL(debouncedXml);
		const settle = (result: ConversionResult) => {
			// A newer request for this format replaces this one
			if (lastRequestedRef.current[kind] !== key) return;
			setConversions((prev) => ({ ...prev, [kind]: result }));
		};
		request.then(
			(code) => settle({ key, code, error: null }),
			(err: unknown) => settle({ key, code: "", error: err instanceof Error ? err.message : String(err) }),
		);
	}, [debouncedXml, fetchXml, selectedFormat, connected, environmentUrl]);

	// The selected format's conversion, if it's for the query on screen; until then it's loading
	const conversionKind = serverConversionFor(selectedFormat);
	const currentConversion = conversionKind ? conversions[conversionKind] : undefined;
	const conversionIsCurrent = currentConversion?.key === conversionKey(fetchXml, environmentUrl);
	const conversionLoading = !!conversionKind && !!fetchXml && connected && !conversionIsCurrent;

	const handleTabSelect = (_e: SelectTabEvent, data: SelectTabData) => {
		setSelectedFormat(data.value as CodeFormat);
	};

	// Resolve the code to display for the current tab
	const currentCode = useCallback((): string => {
		switch (selectedFormat) {
			case "csharp-queryexpr":
			case "sql":
				return conversionIsCurrent ? (currentConversion?.code ?? "") : "";
			default:
				return generateFxsSyncCode(
					selectedFormat as FxsSyncFormat,
					fetchXml,
					selectedFormat === "webapi" ? paEntitySetName : undefined,
					environmentUrl,
				);
		}
		// paEntitySetName loads after the query changes: without it the Web API snippet kept the old
		// (or no) entity set name (#65)
	}, [selectedFormat, fetchXml, conversionIsCurrent, currentConversion, environmentUrl, paEntitySetName]);

	const handleCopy = useCallback(() => {
		const code = currentCode();
		if (!code) return;
		void copyText(code).then((ok) => {
			if (!ok) return;
			setCopied(true);
			setTimeout(() => setCopied(false), 1500);
		});
	}, [currentCode]);

	const isAsyncLoading = conversionLoading;

	const asyncError = conversionIsCurrent ? (currentConversion?.error ?? null) : null;

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
						<Spinner size="medium" label="Generating…" />
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

/** `value`, once it has stopped changing for `delayMs` */
function useDebouncedValue<T>(value: T, delayMs: number): T {
	const [debounced, setDebounced] = useState(value);
	useEffect(() => {
		const timer = setTimeout(() => setDebounced(value), delayMs);
		return () => clearTimeout(timer);
	}, [value, delayMs]);
	return debounced;
}
