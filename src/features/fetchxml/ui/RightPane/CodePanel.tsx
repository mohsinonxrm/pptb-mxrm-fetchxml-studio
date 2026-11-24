/**
 * The Code tab (#233): the query as code in seven languages, on a row of language tabs (Fluent
 * Overflow: what doesn't fit moves into More) with Copy pinned at its end.
 *
 * C# FetchExpression, Web API, Xrm.WebApi, Power Automate and pac CLI are generated in the tool. C#
 * QueryExpression and SQL come from Dataverse (FetchXmlToQueryExpression, FetchXMLToSQL), requested
 * only while their tab is selected, after typing pauses, and once per query and environment (#199).
 * The language picked stays for the session (the parent keeps it); the tab opens on C#
 * QueryExpression, or on C# FetchExpression while not connected.
 */

import { useState, useEffect, useCallback, useRef } from "react";
import Editor from "@monaco-editor/react";
import {
	Button,
	Caption1,
	Link,
	Spinner,
	MessageBar,
	MessageBarActions,
	MessageBarBody,
	Text,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { BarTitle } from "../../../../shared/components/BarTitle";
import {
	ArrowClockwise20Regular,
	Copy20Filled,
	Copy20Regular,
	CopySelect20Regular,
	DocumentDismiss48Regular,
	PlugDisconnected48Regular,
	bundleIcon,
} from "@fluentui/react-icons";
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
	createConversionCache,
	needsConversion,
	serverConversionFor,
	type ConversionCache,
	type ConversionResult,
} from "./serverConversions";
import {
	CODE_LANGUAGES,
	CODE_LANGUAGE_LABELS,
	copyAllText,
	copyDisabledReason,
	needsConnection,
	shownLanguage,
	sourceNote,
	type CodeLanguage,
} from "./codePanelRules";

export type { CodeLanguage } from "./codePanelRules";

const CopyIcon = bundleIcon(Copy20Filled, Copy20Regular);
/** Copy (or Copy all) and the toolbar's padding: the tabs take the rest of the row */
const COPY_WIDTH = 120;

/** Microsoft Learn on Dataverse's SQL: FetchXMLToSQL itself has no reference page */
const SQL_LEARN_MORE =
	"https://learn.microsoft.com/power-apps/developer/data-platform/dataverse-sql-query";

const LANG: Record<CodeLanguage, string> = {
	...FXS_SYNC_FORMAT_LANG,
	"csharp-queryexpr": "csharp",
	sql: "sql",
};

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
		gap: tokens.spacingHorizontalS,
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
		paddingInline: tokens.spacingHorizontalS,
		paddingBlock: tokens.spacingVerticalXXS,
		flexShrink: 0,
	},
	// The language tabs never wrap: the ones that don't fit move into More (#232)
	tabs: { flex: "1 1 0", minWidth: 0 },
	// Where the code comes from, when there's room: it gives way first
	source: {
		flex: "0 1 auto",
		minWidth: 0,
		maxWidth: "30%",
		overflow: "hidden",
		whiteSpace: "nowrap",
		textOverflow: "ellipsis",
		color: tokens.colorNeutralForeground3,
	},
	copy: { flex: `0 0 ${COPY_WIDTH}px` },
	body: {
		flex: 1,
		overflow: "hidden",
		position: "relative",
		display: "flex",
		flexDirection: "column",
	},
	editor: { flex: 1, minHeight: 0 },
	scroll: { flex: 1, minHeight: 0, overflowY: "auto", display: "flex", flexDirection: "column" },
	paPane: {
		flex: 1,
		overflowY: "auto",
		padding: tokens.spacingHorizontalL,
		boxSizing: "border-box",
	},
	bar: { padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS} 0` },
	empty: {
		flex: 1,
		display: "flex",
		flexDirection: "column",
		alignItems: "center",
		justifyContent: "center",
		gap: tokens.spacingVerticalS,
		padding: tokens.spacingHorizontalXXL,
		textAlign: "center",
		color: tokens.colorNeutralForeground3,
	},
	emptyTitle: { color: tokens.colorNeutralForeground1 },
});

interface CodePanelProps {
	fetchXml: string;
	/** The language picked this session, or null for the default (the parent keeps it) */
	language: CodeLanguage | null;
	onLanguageChange: (language: CodeLanguage) => void;
	/**
	 * The conversions asked for and their results, kept by the parent so leaving the tab doesn't
	 * throw them away (#199). Without it the panel keeps its own, for as long as it's mounted.
	 */
	cache?: ConversionCache;
}

export function CodePanel({
	fetchXml,
	language,
	onLanguageChange,
	cache: parentCache,
}: CodePanelProps) {
	const styles = useStyles();
	const { isDark } = useTheme();
	const { connected, environmentUrl } = usePptbContext();
	const selected = shownLanguage(language, connected);
	const label = CODE_LANGUAGE_LABELS[selected];
	const copy = useCopy();

	// Dataverse conversions (C# QueryExpression, SQL): for the selected language only, after typing
	// pauses, and once per query and environment, also across visits to the tab (#199). Try again
	// asks for a failed one again.
	const [ownCache] = useState(createConversionCache);
	const cache = parentCache ?? ownCache;
	const [conversions, setConversions] = useState(() => cache.results);
	const [retries, setRetries] = useState(0);
	// A result that lands while the tab is closed is in the cache; this tab hears about it while open
	useEffect(() => {
		cache.onChange = setConversions;
		setConversions(cache.results);
		return () => {
			if (cache.onChange === setConversions) cache.onChange = undefined;
		};
	}, [cache]);
	const debouncedXml = useDebouncedValue(fetchXml, CONVERSION_DELAY_MS);
	const tabsRef = useRef<HTMLDivElement>(null);
	// Try again: focus stays on the selected language while the error goes away (it unmounts with its
	// button), and the result, or the error again, shows below it
	const tryAgain = () => {
		tabsRef.current?.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')?.focus();
		setRetries((n) => n + 1);
	};

	// Power Automate and Web API: the entity set name from metadata
	const [paEntitySetName, setPaEntitySetName] = useState<string | undefined>(undefined);
	useEffect(() => {
		if (!fetchXml || !connected) {
			setPaEntitySetName(undefined);
			return;
		}
		let cancelled = false;
		const entityNameMatch = fetchXml.match(/<entity\b[^>]*?\sname=["']([^"']+)["']/i);
		if (entityNameMatch) {
			loadEntityMetadata(entityNameMatch[1])
				.then((meta) => {
					if (!cancelled) setPaEntitySetName(meta.EntitySetName);
				})
				.catch(() => {
					/* the generators fall back to a guess */
				});
		}
		return () => {
			cancelled = true;
		};
	}, [fetchXml, connected]);

	useEffect(() => {
		const kind = serverConversionFor(selected);
		// Wait until typing pauses: until then the debounced query is behind the one on screen
		if (!kind || !debouncedXml || debouncedXml !== fetchXml || !connected) return;
		const key = conversionKey(debouncedXml, environmentUrl);
		if (!needsConversion(key, cache.requested[kind], cache.results[kind])) return;
		cache.requested[kind] = key;
		const publish = (results: ConversionCache["results"]) => {
			cache.results = results;
			cache.onChange?.(results);
		};
		// Trying a failed conversion again: show it loading, not the old error
		if (cache.results[kind]?.key === key) publish({ ...cache.results, [kind]: undefined });

		const request =
			kind === "queryexpression"
				? fetchXmlToQueryExpression(debouncedXml).then((qe) =>
						generateCSharpQueryExpression(qe as unknown as QeQueryExpression),
					)
				: fetchXmlToSQL(debouncedXml);
		const settle = (result: ConversionResult) => {
			// A newer request for this language replaces this one
			if (cache.requested[kind] !== key) return;
			publish({ ...cache.results, [kind]: result });
		};
		request.then(
			(code) => settle({ key, code, error: null }),
			(err: unknown) =>
				settle({ key, code: "", error: err instanceof Error ? err.message : String(err) }),
		);
	}, [debouncedXml, fetchXml, selected, connected, environmentUrl, retries, cache]);

	// The selected language's conversion, if it's for the query on screen; until then it's loading
	const conversionKind = serverConversionFor(selected);
	const currentConversion = conversionKind ? conversions[conversionKind] : undefined;
	const conversionIsCurrent = currentConversion?.key === conversionKey(fetchXml, environmentUrl);
	const offline = needsConnection(selected) && !connected;
	const loading = !!conversionKind && !!fetchXml && connected && !conversionIsCurrent;
	const error = conversionIsCurrent ? (currentConversion?.error ?? null) : null;

	const paSpec =
		selected === "powerautomate" ? generatePowerAutomateSpec(fetchXml, paEntitySetName) : null;

	// What Copy copies: the code, or Power Automate's filled-in fields as one block
	const copyText = useCallback((): string => {
		if (paSpec) return copyAllText(paSpec);
		if (conversionKind) return conversionIsCurrent ? (currentConversion?.code ?? "") : "";
		return generateFxsSyncCode(
			selected as FxsSyncFormat,
			fetchXml,
			// paEntitySetName loads after the query changes: without it the Web API snippet kept the old
			// (or no) entity set name (#65)
			selected === "webapi" ? paEntitySetName : undefined,
			environmentUrl,
		);
	}, [
		paSpec,
		conversionKind,
		conversionIsCurrent,
		currentConversion,
		selected,
		fetchXml,
		paEntitySetName,
		environmentUrl,
	]);
	const text = copyText();
	const copyReason = copyDisabledReason(selected, {
		connected,
		loading,
		failed: !!error,
		text,
	});
	// "Copied C# QueryExpression", "Copied all fields" or "Couldn't copy to the clipboard" (#277)
	const handleCopy = () => {
		if (text) void copy(text, paSpec ? "all fields" : label);
	};

	// Text only: no badges or icons with Tooltips inside the tabs, which keyboard users can't reach
	const languageTabs: BarTab[] = CODE_LANGUAGES.map((lang) => ({
		value: lang,
		label: CODE_LANGUAGE_LABELS[lang],
	}));
	const source = sourceNote(selected);

	let body;
	if (offline) {
		body = (
			<div className={styles.empty}>
				<PlugDisconnected48Regular />
				<Text weight="semibold" className={styles.emptyTitle}>
					Not connected
				</Text>
				<Text>Connect to a Dataverse environment to generate {label}.</Text>
			</div>
		);
	} else if (loading) {
		body = (
			<div className={styles.empty}>
				<Spinner size="medium" label="Generating…" />
			</div>
		);
	} else if (error) {
		body = (
			// A long error scrolls, so Try again and Copy details stay reachable
			<div className={styles.scroll}>
				<div className={styles.bar}>
					<MessageBar intent="error" layout="multiline">
						<MessageBarBody>
							<BarTitle>Couldn&apos;t generate {label}</BarTitle>
							{error}
						</MessageBarBody>
						<MessageBarActions>
							<Button
								size="small"
								icon={<ArrowClockwise20Regular />}
								onClick={tryAgain}
							>
								Try again
							</Button>
							<Button
								size="small"
								icon={<CopySelect20Regular />}
								onClick={() => void copy(error, "the error details")}
							>
								Copy details
							</Button>
						</MessageBarActions>
					</MessageBar>
				</div>
				<div className={styles.empty}>
					<DocumentDismiss48Regular />
					<Text weight="semibold" className={styles.emptyTitle}>
						No code to show
					</Text>
					<Text>Fix the problem above, then try again.</Text>
				</div>
			</div>
		);
	} else if (paSpec) {
		body = (
			<div className={styles.paPane}>
				<PowerAutomatePane spec={paSpec} />
			</div>
		);
	} else {
		body = (
			<>
				<div className={styles.editor}>
					<Editor
						height="100%"
						language={LANG[selected]}
						value={text}
						theme={isDark ? "vs-dark" : "light"}
						options={{
							readOnly: true,
							minimap: { enabled: false },
							scrollBeyondLastLine: false,
							wordWrap: selected === "webapi" ? "on" : "off",
							fontSize: 13,
							lineNumbers: "on",
							renderLineHighlight: "none",
							folding: true,
							automaticLayout: true,
						}}
					/>
				</div>
			</>
		);
	}

	return (
		<div className={styles.root}>
			{/* The language tabs, where the code comes from, then Copy (#232, #233) */}
			<div className={styles.header}>
				<div className={styles.tabs} ref={tabsRef}>
					<OverflowTabList
						ariaLabel="Code languages"
						size="small"
						appearance="subtle"
						tabs={languageTabs}
						selectedValue={selected}
						onTabSelect={(value) => onLanguageChange(value as CodeLanguage)}
					/>
				</div>
				{source && (
					<Caption1 className={styles.source} title={source}>
						{source}
					</Caption1>
				)}
				<div className={styles.copy}>
					<OverflowToolbar
						ariaLabel="Code commands"
						size="small"
						align="end"
						commands={[
							{
								id: "copy",
								label: paSpec ? "Copy all" : "Copy",
								icon: CopyIcon,
								kind: "button",
								pinned: true,
								hint: paSpec
									? "Copy every filled-in field as one block"
									: `Copy the ${label} code`,
								disabledReason: copyReason,
								onClick: handleCopy,
							},
						]}
					/>
				</div>
			</div>
			<div className={styles.body}>
				{/* Above the states, so they don't come and go (and get announced again) while loading */}
				{selected === "sql" && !offline && (
					<div className={styles.bar}>
						<MessageBar intent="info">
							<MessageBarBody>
								<BarTitle>SQL is a preview</BarTitle>
								It comes from FetchXMLToSQL, which isn&apos;t documented and can
								change.{" "}
								<Link href={SQL_LEARN_MORE} target="_blank" rel="noreferrer" inline>
									Learn more
								</Link>
							</MessageBarBody>
						</MessageBar>
					</div>
				)}
				{language === null && !connected && (
					<div className={styles.bar}>
						<MessageBar intent="info">
							<MessageBarBody>
								Connect to an environment to generate QueryExpression.
							</MessageBarBody>
						</MessageBar>
					</div>
				)}
				{body}
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
