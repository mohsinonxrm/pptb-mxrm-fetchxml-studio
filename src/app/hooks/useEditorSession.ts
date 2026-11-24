/**
 * The FetchXML editor's session and Run (#278): editor mode and its text, loading the editor's
 * FetchXML into the tree before an action (#102, #89), Run and "Run anyway", and which right-pane
 * tab is open. It lived in PreviewTabs; it's lifted to AppShell so the app bar's Run, Save and Send
 * to, which sit outside the right pane, use the same editor sync (EditorSyncContext) as the tabs.
 */
import { useCallback, useEffect, useMemo, useRef, useState, type MutableRefObject } from "react";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import type { ParseResult, ParseWarning } from "../../features/fetchxml/model/fetchxmlParser";
import {
	validateFetchXmlSyntax,
	parseFetchXml,
	getBlockingWarnings,
	getExecutionErrors,
	getParseToTreeChanges,
} from "../../features/fetchxml/model/fetchxmlParser";
import { generateFetchXml } from "../../features/fetchxml/model/fetchxml";
import type { LoadFetchXmlOptions } from "../../features/fetchxml/state/builderContext";
import type {
	EditorSyncAction,
	EditorSyncResult,
	EditorSyncValue,
} from "../../features/fetchxml/ui/RightPane/editorSync";
import {
	invalidProblem,
	mightRejectProblem,
	syntaxProblem,
	type EditorProblem,
	type EditorProblemKey,
} from "../../features/fetchxml/ui/RightPane/editorProblems";
import type { FetchXmlEditorHandle } from "../../features/fetchxml/ui/RightPane/FetchXmlEditor";
import type { PreviewTab } from "../../features/fetchxml/ui/RightPane/previewHeaderRules";
import { NO_QUERY_REASON } from "../../features/fetchxml/ui/Toolbar/useSaveViewCommand";

/** Why Run is off with an empty editor in editor mode */
export const EMPTY_EDITOR_REASON = "Enter FetchXML in the editor first";
/** Why Run is off while the query runs or more rows load */
export const RUNNING_REASON = "Running the query…";

export interface EditorSessionInputs {
	/** The tree's FetchXML */
	xml: string;
	onParseToTree?: (xmlString: string, options?: LoadFetchXmlOptions) => ParseResult;
	/**
	 * Run the tree's query. After an editor-mode sync (#102) the new tree and its XML are passed
	 * along, because builder state updates arrive on the next render.
	 */
	onRun?: (synced?: { xml: string; fetchQuery: FetchNode }) => void;
	/** Why the tree's query can't run, e.g. a quick find view without a test value (#142) */
	runDisabledReason?: string;
	isRunning?: boolean;
	isLoadingMore?: boolean;
	/** Ctrl+Enter in the editor while Run can't act: the reason, to announce (#186 PR 3) */
	onRunBlocked?: (reason: string) => void;
}

export interface PendingSync {
	action: EditorSyncAction;
	changes: ParseWarning[];
	resolve: (confirmed: boolean) => void;
}

export interface EditorSession {
	selectedTab: PreviewTab;
	setSelectedTab: (tab: PreviewTab) => void;
	isEditorModeActive: boolean;
	/** Editor mode is on and the editor has text */
	editorHasContent: boolean;
	/** The editor's text a moment after typing stops, for the Advisories tab (#97) */
	editorText: string;
	/** The editor's text now, with the debounce flushed (Check with Dataverse checks this) */
	flushEditorText: () => string;
	fetchXmlEditorRef: MutableRefObject<FetchXmlEditorHandle | null>;
	/** Cleaned FetchXML to show in the editor after a sync changed the query (#102) */
	editorReplace: { xml: string; id: number } | null;
	/** A sync waiting for the user to confirm the changes loading into the tree makes (#89) */
	pendingSync: PendingSync | null;
	resolvePendingSync: (confirmed: boolean) => void;
	onEditorStateChange: (isActive: boolean, editorXml: string) => void;
	/** For EditorSyncContext, around the app bar and the right pane */
	syncValue: EditorSyncValue;
	/** Run: the editor's FetchXML in editor mode (after its checks), otherwise the tree's */
	run: () => void;
	/** Why Run can't be pressed now, if it can't */
	runDisabledReason: string | undefined;
	/** The query is running, or more rows are loading: Run shows it's busy */
	isRunning: boolean;
	/** The FetchXML the user sees: the editor's in editor mode, otherwise the tree's */
	getCurrentXml: () => string;
	/** For FetchXmlEditor's resetRequest: bumped by resetEditor and discardEditor */
	editorResetId: number;
	/** New (#278): leave editor mode, drop the editor's text and open the FetchXML tab */
	resetEditor: () => void;
	/** Open view (#278): leave editor mode and drop the editor's text, on whichever tab is open */
	discardEditor: () => void;
}

export function useEditorSession({
	xml,
	onParseToTree,
	onRun,
	runDisabledReason: treeRunReason,
	isRunning = false,
	isLoadingMore = false,
	onRunBlocked,
}: EditorSessionInputs): EditorSession {
	const [selectedTab, setSelectedTab] = useState<PreviewTab>("xml");
	// isEditorModeActive in state controls renders; the ref holds the latest XML without a render
	// on each keystroke
	const [isEditorModeActive, setIsEditorModeActive] = useState(false);
	const editorXmlRef = useRef<string>("");
	// The FetchXML editor owns the bars about its FetchXML; the checks here report to it (#277)
	const fetchXmlEditorRef = useRef<FetchXmlEditorHandle | null>(null);
	// The editor FetchXML that "Run anyway" was offered for: it only applies to that text (#89)
	const runAnywayXmlRef = useRef<string | null>(null);
	/** Reports a check's findings to the editor; findings switch to the FetchXML tab, where they are */
	const reportToEditor = useCallback((checked: EditorProblemKey[], found: EditorProblem[]) => {
		fetchXmlEditorRef.current?.report(checked, found);
		if (found.length > 0) setSelectedTab("xml");
	}, []);

	// Whether the editor buffer has any content (Run is enabled on it in editor mode, #102)
	const [editorHasContent, setEditorHasContent] = useState(false);
	const [editorReplace, setEditorReplace] = useState<{ xml: string; id: number } | null>(null);
	const [pendingSync, setPendingSync] = useState<PendingSync | null>(null);
	const [editorResetId, setEditorResetId] = useState(0);
	const discardEditor = useCallback(() => {
		runAnywayXmlRef.current = null;
		setEditorResetId((n) => n + 1);
	}, []);
	const resetEditor = useCallback(() => {
		discardEditor();
		setSelectedTab("xml");
	}, [discardEditor]);

	const getCurrentXml = useCallback(() => {
		return isEditorModeActive ? editorXmlRef.current || xml : xml;
	}, [isEditorModeActive, xml]);

	// The editor's text for the Advisories tab (#97), a moment after typing stops
	const [editorText, setEditorText] = useState("");
	const editorTextTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
	useEffect(() => () => clearTimeout(editorTextTimer.current), []);
	const flushEditorText = useCallback(() => {
		const text = editorXmlRef.current;
		clearTimeout(editorTextTimer.current);
		setEditorText(text);
		return text;
	}, []);

	const onEditorStateChange = useCallback((isActive: boolean, editorXml: string) => {
		editorXmlRef.current = editorXml;
		setIsEditorModeActive((prev) => (prev !== isActive ? isActive : prev));
		setEditorHasContent(isActive && editorXml.trim().length > 0);
		// A pending "Run anyway" only applies to the exact XML it warned about
		if (
			runAnywayXmlRef.current !== null &&
			(!isActive || runAnywayXmlRef.current !== editorXml)
		) {
			runAnywayXmlRef.current = null;
			// The text changed: Run's warning was about the old text. Leaving editor mode needs
			// nothing here: turning it off clears the bars, and Parse to tree reports its own.
			if (isActive) fetchXmlEditorRef.current?.drop("warning", "execute");
		}
		clearTimeout(editorTextTimer.current);
		editorTextTimer.current = setTimeout(() => setEditorText(isActive ? editorXml : ""), 300);
	}, []);

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
				reportToEditor(
					["error"],
					[syntaxProblem(syntaxResult.error || "Invalid XML syntax")],
				);
				return { status: "cancelled" };
			}
			const preview = parseFetchXml(editorXml);
			if (!preview.success || !preview.fetchNode) {
				reportToEditor(["error"], [invalidProblem(preview.errors)]);
				return { status: "cancelled" };
			}
			// Known to be rejected by Dataverse: block (Parse to Tree can often correct these)
			const executionErrors = getExecutionErrors(preview);
			if (executionErrors.length > 0) {
				reportToEditor(["error"], [invalidProblem(executionErrors)]);
				return { status: "cancelled" };
			}
			// No errors now: an earlier error bar is out of date
			reportToEditor(["error"], []);
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
				reportToEditor(
					["error"],
					[
						invalidProblem(
							result?.errors ?? [
								{ message: "The FetchXML couldn't be loaded into the tree." },
							],
						),
					],
				);
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
		[isEditorModeActive, xml, onParseToTree, reportToEditor],
	);

	// Ctrl+Enter in the editor (#186) runs as Run does, through the latest run and reason
	const runShortcut = useRef<() => void>(() => {});
	const runFromEditor = useCallback(() => runShortcut.current(), []);
	const syncValue = useMemo(
		() => ({
			sync: syncEditorToTree,
			editorHasContent: isEditorModeActive && editorHasContent,
			runFromEditor,
		}),
		[syncEditorToTree, isEditorModeActive, editorHasContent, runFromEditor],
	);

	const resolvePendingSync = (confirmed: boolean) => {
		pendingSync?.resolve(confirmed);
		setPendingSync(null);
	};

	/** Sync the editor into the tree (editor mode), then run the tree's query. */
	const runQuery = async () => {
		const synced = await syncEditorToTree("execute");
		if (synced.status === "cancelled") return;
		setSelectedTab("results");
		onRun?.(
			synced.status === "synced"
				? { xml: synced.fetchXml, fetchQuery: synced.fetchQuery }
				: undefined,
		);
	};

	// The bar's Run anyway runs the latest version, not the one from when the bar was shown
	const runQueryRef = useRef(runQuery);
	useEffect(() => {
		runQueryRef.current = runQuery;
	});

	/** Run the editor XML despite warnings; Dataverse's own error is shown if it fails. */
	const runAnyway = () => {
		if (runAnywayXmlRef.current === null) return;
		runAnywayXmlRef.current = null;
		fetchXmlEditorRef.current?.report(["warning"], []);
		void runQueryRef.current();
	};
	const runAnywayRef = useRef(runAnyway);
	useEffect(() => {
		runAnywayRef.current = runAnyway;
	});

	const run = () => {
		if (isEditorModeActive) {
			// Can't be determined client-side: ask before sending. Errors are checked by the sync.
			const editorXml = editorXmlRef.current;
			const preview = parseFetchXml(editorXml);
			if (preview.success && getExecutionErrors(preview).length === 0) {
				const blocking = getBlockingWarnings(preview);
				if (blocking.length > 0) {
					runAnywayXmlRef.current = editorXml;
					// It parses and has no errors, so an error bar from before is out of date too
					reportToEditor(
						["error", "warning"],
						[
							mightRejectProblem(blocking, {
								action: {
									label: "Run anyway",
									onClick: () => runAnywayRef.current(),
								},
								source: "execute",
							}),
						],
					);
					return;
				}
				// Nothing Dataverse might reject now: a warning from before is out of date
				runAnywayXmlRef.current = null;
				reportToEditor(["warning"], []);
			}
		}
		void runQuery();
	};

	// The quick find reason follows the tree's query; in editor mode the editor's query runs after
	// syncing, and Run stops there if it still needs a test value (#142)
	const runDisabledReason =
		isRunning || isLoadingMore
			? RUNNING_REASON
			: ((isEditorModeActive ? undefined : treeRunReason) ??
				(isEditorModeActive
					? editorHasContent
						? undefined
						: EMPTY_EDITOR_REASON
					: xml
						? undefined
						: NO_QUERY_REASON));
	useEffect(() => {
		runShortcut.current = () => {
			if (runDisabledReason) onRunBlocked?.(runDisabledReason);
			else run();
		};
	});

	return {
		selectedTab,
		setSelectedTab,
		isEditorModeActive,
		editorHasContent,
		editorText,
		flushEditorText,
		fetchXmlEditorRef,
		editorReplace,
		pendingSync,
		resolvePendingSync,
		onEditorStateChange,
		syncValue,
		run,
		runDisabledReason,
		isRunning: isRunning || isLoadingMore,
		getCurrentXml,
		editorResetId,
		resetEditor,
		discardEditor,
	};
}
