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
	Spinner,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Badge,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Copy20Filled, Copy20Regular, Info20Regular, bundleIcon } from "@fluentui/react-icons";
import {
	OverflowTabList,
	OverflowToolbar,
	type BarTab,
} from "../../../../shared/components/commandBar";
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
import { useCopy } from "../../../../shared/feedback/useCopy";
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

const CopyIcon = bundleIcon(Copy20Filled, Copy20Regular);
/** Copy and the toolbar's padding: the tabs take the rest of the row */
const COPY_WIDTH = 104;

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
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
		paddingInline: tokens.spacingHorizontalS,
		paddingBlock: tokens.spacingVerticalXXS,
		flexShrink: 0,
		backgroundColor: tokens.colorNeutralBackground2,
	},
	// The language tabs never wrap: the ones that don't fit move into More (#232)
	tabs: { flex: "1 1 0", minWidth: 0 },
	copy: { flex: `0 0 ${COPY_WIDTH}px` },
	tabExtras: {
		display: "inline-flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalXS,
		marginInlineStart: tokens.spacingHorizontalXS,
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
	const copy = useCopy();

	// Dataverse conversions (C# QueryExpression, SQL): for the selected format only, after typing
	// pauses, and once per query and environment (#199)
	const [conversions, setConversions] = useState<
		Partial<Record<ServerConversion, ConversionResult>>
	>({});
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
		if (!needsConversion(key, lastRequestedRef.current[kind], conversionsRef.current[kind]))
			return;
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
			(err: unknown) =>
				settle({ key, code: "", error: err instanceof Error ? err.message : String(err) }),
		);
	}, [debouncedXml, fetchXml, selectedFormat, connected, environmentUrl]);

	// The selected format's conversion, if it's for the query on screen; until then it's loading
	const conversionKind = serverConversionFor(selectedFormat);
	const currentConversion = conversionKind ? conversions[conversionKind] : undefined;
	const conversionIsCurrent = currentConversion?.key === conversionKey(fetchXml, environmentUrl);
	const conversionLoading = !!conversionKind && !!fetchXml && connected && !conversionIsCurrent;

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
	}, [
		selectedFormat,
		fetchXml,
		conversionIsCurrent,
		currentConversion,
		environmentUrl,
		paEntitySetName,
	]);

	const handleCopy = useCallback(() => {
		const code = currentCode();
		if (!code) return;
		// "Copied C# QueryExpression" or "Couldn't copy to the clipboard" (#277)
		void copy(code, FORMAT_LABELS[selectedFormat]);
	}, [currentCode, copy, selectedFormat]);

	const isAsyncLoading = conversionLoading;

	const asyncError = conversionIsCurrent ? (currentConversion?.error ?? null) : null;

	const showOffline =
		(selectedFormat === "csharp-queryexpr" || selectedFormat === "sql") && !connected;

	const code = currentCode();
	const monacoTheme = isDark ? "vs-dark" : "light";

	// Why Copy is off: Power Automate has its own per-field Copy buttons
	const copyDisabledReason =
		selectedFormat === "powerautomate"
			? "Use the per-field Copy buttons below"
			: showOffline
				? "Connect to an environment to generate"
				: isAsyncLoading
					? "Generating…"
					: !code
						? "There's no code to copy yet"
						: undefined;

	// Each language, with Preview (SQL) and the not-connected hint after its name
	const languageTabs: BarTab[] = ALL_FORMATS.map((fmt) => {
		const preview = fmt === "sql";
		const offline = (fmt === "csharp-queryexpr" || fmt === "sql") && !connected;
		return {
			value: fmt,
			label: FORMAT_LABELS[fmt],
			content:
				preview || offline ? (
					<span className={styles.tabExtras}>
						{preview && (
							<Tooltip
								content="Undocumented API — may change without notice"
								relationship="label"
							>
								<Badge size="small" appearance="tint" color="warning">
									Preview
								</Badge>
							</Tooltip>
						)}
						{offline && (
							<Tooltip
								content="Connect to an environment to generate"
								relationship="label"
							>
								<Info20Regular style={{ opacity: 0.5, width: 14, height: 14 }} />
							</Tooltip>
						)}
					</span>
				) : undefined,
		};
	});
	const monacoLang = FORMAT_LANG[selectedFormat];

	return (
		<div className={styles.root}>
			{/* Header: the language tabs, then Copy (#232) */}
			<div className={styles.header}>
				<div className={styles.tabs}>
					<OverflowTabList
						ariaLabel="Code languages"
						size="small"
						appearance="subtle"
						tabs={languageTabs}
						selectedValue={selectedFormat}
						onTabSelect={(value) => setSelectedFormat(value as CodeFormat)}
					/>
				</div>
				<div className={styles.copy}>
					<OverflowToolbar
						ariaLabel="Code commands"
						size="small"
						align="end"
						commands={[
							{
								id: "copy",
								label: "Copy",
								ariaLabel: "Copy code to clipboard",
								icon: CopyIcon,
								kind: "button",
								pinned: true,
								hint: "Copy to clipboard",
								disabledReason: copyDisabledReason,
								onClick: handleCopy,
							},
						]}
					/>
				</div>
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
							{selectedFormat === "csharp-queryexpr" ? "C# QueryExpression" : "SQL"}{" "}
							code.
						</p>
					</div>
				)}

				{/* Power Automate pane — field-by-field form */}
				{!isAsyncLoading &&
					!asyncError &&
					!showOffline &&
					selectedFormat === "powerautomate" && (
						<div className={styles.paPane}>
							<PowerAutomatePane
								spec={generatePowerAutomateSpec(fetchXml, paEntitySetName)}
							/>
						</div>
					)}

				{/* Monaco editor for all other formats */}
				{!isAsyncLoading &&
					!asyncError &&
					!showOffline &&
					selectedFormat !== "powerautomate" && (
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
