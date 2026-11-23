/**
 * Monaco editor for FetchXML with Editor Mode toggle
 * - Read-only mode: Shows generated XML with Copy button
 * - Editor mode: Editable with Paste, Load File, Reset, Validate, Parse to tree buttons
 */

import {
	useState,
	useRef,
	useCallback,
	useEffect,
	forwardRef,
	useImperativeHandle,
	type ForwardedRef,
} from "react";
import Editor from "@monaco-editor/react";
import type * as MonacoEditor from "monaco-editor/editor";
import {
	Button,
	makeStyles,
	Spinner,
	tokens,
	Switch,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	MessageBarActions,
	MessageBarGroup,
} from "@fluentui/react-components";
import {
	Copy20Filled,
	Copy20Regular,
	ClipboardPaste20Filled,
	ClipboardPaste20Regular,
	FolderOpen20Filled,
	FolderOpen20Regular,
	TextIndentIncrease20Filled,
	TextIndentIncrease20Regular,
	ArrowReset20Filled,
	ArrowReset20Regular,
	Checkmark20Filled,
	Checkmark20Regular,
	TreeDeciduous20Filled,
	TreeDeciduous20Regular,
	Dismiss16Regular,
	bundleIcon,
} from "@fluentui/react-icons";
import { OverflowToolbar, type BarCommand } from "../../../../shared/components/commandBar";
import { useTheme } from "../../../../shared/contexts/useTheme";
import {
	parseFetchXml,
	validateFetchXmlSyntax,
	getBlockingWarnings,
	getExecutionErrors,
	getParseToTreeChanges,
} from "../../model/fetchxmlParser";
import type { ParseResult, ParseWarning } from "../../model/fetchxmlParser";
import { generateFetchXml } from "../../model/fetchxml";
import { useNotify } from "../../../../shared/feedback/useNotify";
import { useCopy } from "../../../../shared/feedback/useCopy";
import { ParseIssueList } from "./ParseIssueList";
import {
	applyCheck,
	FULL_CHECK,
	invalidProblem,
	mightRejectProblem,
	notesProblem,
	syntaxProblem,
	type EditorProblem,
	type EditorProblemKey,
} from "./editorProblems";
import { ParseChangesDialog } from "./ParseChangesDialog";
import { attachFetchXmlLanguageFeatures } from "../../model/fetchxmlIntellisense";
import { formatFetchXml } from "../../model/fetchxmlFormatter";
import { formatKeys, getEntry } from "../../../../app/shortcuts/registry";

/** FetchXML as the tree would write it, to compare meaning rather than formatting; null if it doesn't parse. */
function canonicalFetchXml(xmlText: string): string | null {
	const result = parseFetchXml(xmlText);
	return result.success && result.fetchNode ? generateFetchXml(result.fetchNode) : null;
}

const TREE_CHANGED_TITLE = "The tree changed while you were editing";

/**
 * Whether the Advisories tab lists this finding (advisories.ts collectAdvisories): a query-rule
 * finding (they carry a nodeId) that's an error, a warning or advice, not a setting Dataverse ignores.
 */
const listedInAdvisories = (w: ParseWarning) =>
	w.nodeId !== undefined && (w.severity !== "info" || w.kind === "advice");

const CopyIcon = bundleIcon(Copy20Filled, Copy20Regular);
const PasteIcon = bundleIcon(ClipboardPaste20Filled, ClipboardPaste20Regular);
const LoadIcon = bundleIcon(FolderOpen20Filled, FolderOpen20Regular);
const ResetIcon = bundleIcon(ArrowReset20Filled, ArrowReset20Regular);
const ValidateIcon = bundleIcon(Checkmark20Filled, Checkmark20Regular);
const ParseIcon = bundleIcon(TreeDeciduous20Filled, TreeDeciduous20Regular);
const FormatIcon = bundleIcon(TextIndentIncrease20Filled, TextIndentIncrease20Regular);

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		overflow: "hidden",
	},
	toolbar: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		paddingInline: tokens.spacingHorizontalS,
		paddingBlock: tokens.spacingVerticalXXS,
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
		backgroundColor: tokens.colorNeutralBackground2,
		flexShrink: 0,
	},
	commands: { flex: "1 1 0", minWidth: 0 },
	editorWrapper: {
		flex: 1,
		overflow: "hidden",
		minHeight: 0,
	},
	loadingContainer: {
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		height: "100%",
		backgroundColor: tokens.colorNeutralBackground1,
	},
	messageBarContainer: {
		padding: tokens.spacingVerticalS,
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		maxHeight: "150px",
		overflowY: "auto",
		flexShrink: 0,
	},
	switchLabel: {
		fontSize: tokens.fontSizeBase200,
		fontWeight: tokens.fontWeightSemibold,
		whiteSpace: "nowrap",
	},
	// The Switch never shrinks; the commands give way
	modeSwitch: { flexShrink: 0 },
});

/** For PreviewTabs: the editor-mode checks of Execute, Save view and Send to tool report here (#277) */
export interface FetchXmlEditorHandle {
	/** Replaces the bars for the keys `checked` with what the check `found` (see applyCheck) */
	report: (checked: EditorProblemKey[], found: EditorProblem[]) => void;
	/** Removes the bar for `key` only if `source` raised it (a later check may have replaced it) */
	drop: (key: EditorProblemKey, source: NonNullable<EditorProblem["source"]>) => void;
}

interface FetchXmlEditorProps {
	xml: string;
	onParseToTree?: (xmlString: string) => ParseResult;
	/**
	 * Called whenever editor mode or its content changes.
	 * `isActive=false` means editor mode is off (parent should use tree XML).
	 * `isActive=true` means editor mode is on; `xml` is the current editor content.
	 */
	onEditorStateChange?: (isActive: boolean, xml: string) => void;
	/**
	 * Replace the editor buffer (editor mode only), e.g. after an editor-mode action loaded the
	 * FetchXML into the tree and the tree changed it (#102). A new `id` triggers the replacement.
	 */
	replaceRequest?: { xml: string; id: number } | null;
	/** A new value leaves editor mode and drops the editor's text: New started a new query (#278) */
	resetRequest?: number;
	/** Switch to the Advisories tab, from Validate's message bars (#97) */
	onViewAdvisories?: () => void;
}

function FetchXmlEditorImpl(
	{
		xml,
		onParseToTree,
		onEditorStateChange,
		replaceRequest,
		resetRequest,
		onViewAdvisories,
	}: FetchXmlEditorProps,
	ref: ForwardedRef<FetchXmlEditorHandle>,
) {
	const styles = useStyles();
	const { isDark } = useTheme();
	const editorRef = useRef<MonacoEditor.editor.IStandaloneCodeEditor | null>(null);
	const fileInputRef = useRef<HTMLInputElement>(null);

	// Editor mode state
	const [isEditorMode, setIsEditorMode] = useState(false);
	const [editedXml, setEditedXml] = useState("");
	// One bar per condition, most severe first (#277)
	const [messages, setMessages] = useState<EditorProblem[]>([]);
	// Parse to tree awaiting confirmation because it would change the query (#89)
	const [pendingParse, setPendingParse] = useState<{
		xml: string;
		changes: ParseWarning[];
	} | null>(null);

	// Short confirmations go to the app's Toaster; bars here are only for problems in this XML (#277)
	const { notify } = useNotify();
	const copy = useCopy();

	// Get current XML value based on mode
	const currentXml = isEditorMode ? editedXml : xml;

	// A check replaces the bars for what it checked
	// A check replaces the bars for what it checked. A bar updated in place isn't announced again:
	// Fluent announces a MessageBar when it mounts, through the AriaLiveAnnouncer (#186)
	const report = useCallback((checked: EditorProblemKey[], found: EditorProblem[]) => {
		setMessages((prev) => applyCheck(prev, checked, found));
	}, []);
	const drop = useCallback(
		(key: EditorProblemKey, source: NonNullable<EditorProblem["source"]>) => {
			setMessages((prev) => prev.filter((m) => !(m.key === key && m.source === source)));
		},
		[],
	);
	useImperativeHandle(ref, () => ({ report, drop }), [report, drop]);

	const dismissMessage = useCallback((key: EditorProblemKey) => {
		setMessages((prev) => prev.filter((m) => m.key !== key));
	}, []);

	// Clear all messages
	const clearMessages = useCallback(() => {
		setMessages([]);
	}, []);

	// Tree FetchXML the editor last matched. While editor mode is on, tree changes (e.g. edits in the
	// property panel) are followed when the editor still means the same as this; otherwise the editor
	// keeps its text and says so, because Execute / Save / Send would replace the tree with it (#102).
	const lastTreeXmlRef = useRef(xml);
	useEffect(() => {
		if (!isEditorMode) {
			lastTreeXmlRef.current = xml;
			return;
		}
		if (xml === lastTreeXmlRef.current) return;
		const editorMeaning = canonicalFetchXml(editedXml);
		if (editorMeaning === xml) {
			// The tree now matches the editor (e.g. after an editor-mode sync)
			lastTreeXmlRef.current = xml;
			return;
		}
		if (editorMeaning !== null && editorMeaning === lastTreeXmlRef.current) {
			// The editor hasn't changed the query: follow the tree
			lastTreeXmlRef.current = xml;
			setEditedXml(xml);
			onEditorStateChange?.(true, xml);
			return;
		}
		lastTreeXmlRef.current = xml;
		report(
			["tree-changed"],
			[
				{
					key: "tree-changed",
					intent: "warning",
					title: TREE_CHANGED_TITLE,
					body: "Run, Save and Send to tool use the editor's FetchXML and will replace the tree's changes. Use Reset to take the tree's FetchXML instead.",
				},
			],
		);
		// Only tree changes drive this; the editor's own edits are compared when the tree changes
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [xml, isEditorMode]);

	// Show what the tree loaded when an editor-mode action changed the query (#102)
	useEffect(() => {
		if (!replaceRequest || !isEditorMode) return;
		setEditedXml(replaceRequest.xml);
		onEditorStateChange?.(true, replaceRequest.xml);
		notify.info({
			title: "Editor updated",
			body: "The editor now shows the FetchXML that was loaded into the tree, with the changes you confirmed.",
		});
		// Only a new request replaces the buffer
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [replaceRequest?.id]);

	// Toggle editor mode
	const handleToggleEditorMode = useCallback(
		(checked: boolean) => {
			if (checked) {
				// Entering editor mode - copy current XML to edit buffer
				setEditedXml(xml);
				clearMessages();
				onEditorStateChange?.(true, xml);
			} else {
				// The bars were about the editor's FetchXML, which is no longer used (Parse to
				// tree also leaves editor mode, but keeps its results)
				clearMessages();
				onEditorStateChange?.(false, "");
			}
			setIsEditorMode(checked);
		},
		[xml, clearMessages, onEditorStateChange],
	);

	// New (#278): back to the tree's FetchXML, out of editor mode
	useEffect(() => {
		if (!resetRequest || !isEditorMode) return;
		handleToggleEditorMode(false);
		// Only a new request resets
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [resetRequest]);

	// Copy to clipboard
	const handleCopy = useCallback(async () => {
		await copy(currentXml, "FetchXML");
	}, [currentXml, copy]);

	// Paste from clipboard
	const handlePaste = useCallback(async () => {
		try {
			const text = await navigator.clipboard.readText();
			if (text.trim()) {
				setEditedXml(text);
				onEditorStateChange?.(true, text);
				notify.info("Pasted from the clipboard");
			}
		} catch (err) {
			console.error("Failed to paste:", err);
			notify.error("Couldn't read the clipboard");
		}
	}, [notify, onEditorStateChange]);

	// Load from file
	const handleLoadFile = useCallback(() => {
		fileInputRef.current?.click();
	}, []);

	const handleFileChange = useCallback(
		(event: React.ChangeEvent<HTMLInputElement>) => {
			const file = event.target.files?.[0];
			if (file) {
				const reader = new FileReader();
				reader.onload = (e) => {
					const content = e.target?.result as string;
					setEditedXml(content);
					onEditorStateChange?.(true, content);
					notify.info(`Loaded ${file.name}`);
				};
				reader.onerror = () => {
					notify.error({ title: "Couldn't read the file", body: file.name });
				};
				reader.readAsText(file);
			}
			// Reset input so same file can be loaded again
			event.target.value = "";
		},
		[notify, onEditorStateChange],
	);

	// Reset to tree-generated XML
	const handleReset = useCallback(() => {
		setEditedXml(xml);
		onEditorStateChange?.(true, xml);
		clearMessages();
		notify.info("Reverted to the FetchXML from the tree");
	}, [xml, clearMessages, notify, onEditorStateChange]);

	// Format (#186): re-indent the editor's FetchXML with the safe formatter, as one step of the
	// editor's own undo (Ctrl+Z takes it back). Shift+Alt+F does the same through Monaco.
	const handleFormat = useCallback(() => {
		const editor = editorRef.current;
		const model = editor?.getModel();
		const text = model ? model.getValue() : editedXml;
		const result = formatFetchXml(text);
		if (!result.ok) {
			notify.warning(result.error);
			return;
		}
		if (result.text === text) return;
		if (editor && model) {
			editor.pushUndoStop();
			editor.executeEdits("format", [
				{ range: model.getFullModelRange(), text: result.text },
			]);
			editor.pushUndoStop();
		} else {
			setEditedXml(result.text);
		}
	}, [editedXml, notify]);

	// Validate XML
	const handleValidate = useCallback(() => {
		// First check syntax
		const syntaxResult = validateFetchXmlSyntax(editedXml);
		if (!syntaxResult.valid) {
			report(FULL_CHECK, [syntaxProblem(syntaxResult.error || "Invalid XML syntax")]);
			return;
		}

		// Then do a full parse to classify findings
		const parseResult = parseFetchXml(editedXml);
		if (!parseResult.success) {
			report(FULL_CHECK, [invalidProblem(parseResult.errors)]);
			return;
		}

		// One bar per level. A bar offers "View all in Advisories" when the tab lists one of its
		// findings (#97)
		const notes = parseResult.warnings.filter((w) => w.severity === "info");
		const blocking = getBlockingWarnings(parseResult);
		const executionErrors = getExecutionErrors(parseResult);
		const found: EditorProblem[] = [];
		if (executionErrors.length > 0)
			found.push(invalidProblem(executionErrors, executionErrors.some(listedInAdvisories)));
		if (blocking.length > 0)
			found.push(
				mightRejectProblem(blocking, { advisories: blocking.some(listedInAdvisories) }),
			);
		if (notes.length > 0) found.push(notesProblem(notes, notes.some(listedInAdvisories)));
		report(FULL_CHECK, found);
		if (blocking.length === 0 && executionErrors.length === 0) {
			notify.success({ title: "Valid FetchXML", body: "Dataverse should accept it." });
		}
	}, [editedXml, report, notify]);

	// Load XML into the tree. `changesConfirmed` is true when the user already reviewed the
	// changes in the confirmation dialog, so they aren't repeated as notes.
	const commitParseToTree = useCallback(
		(xmlToParse: string, changesConfirmed: boolean) => {
			if (!onParseToTree) return;

			const result = onParseToTree(xmlToParse);

			if (!result.success) {
				report(FULL_CHECK, [invalidProblem(result.errors)]);
				return;
			}

			const changes = new Set(getParseToTreeChanges(result));
			// Kept as written by Parse to tree, so still a problem in the tree
			const stillInvalid = result.warnings.filter((w) => w.kind === "invalid");
			const stillErrors = stillInvalid.filter((w) => w.severity === "error");
			const stillWarnings = stillInvalid.filter((w) => w.severity === "warning");
			const notes = result.warnings.filter(
				(w) => w.severity === "info" && !(changesConfirmed && changes.has(w)),
			);
			const found: EditorProblem[] = [];
			if (stillErrors.length > 0) found.push(invalidProblem(stillErrors));
			if (stillWarnings.length > 0) found.push(mightRejectProblem(stillWarnings));
			if (notes.length > 0) found.push(notesProblem(notes));
			// The tree now holds what the editor had: the tree-changed bar is out of date too
			report([...FULL_CHECK, "tree-changed"], found);
			notify.success("Loaded into the tree");

			// Exit editor mode after successful parse
			onEditorStateChange?.(false, "");
			setIsEditorMode(false);
		},
		[onParseToTree, report, notify, onEditorStateChange],
	);

	// Parse to tree — confirm first when the tree would change the query
	const handleParseToTree = useCallback(() => {
		if (!onParseToTree) return;

		const preview = parseFetchXml(editedXml);
		if (!preview.success) {
			report(FULL_CHECK, [invalidProblem(preview.errors)]);
			return;
		}

		const changes = getParseToTreeChanges(preview);
		if (changes.length > 0) {
			setPendingParse({ xml: editedXml, changes });
			return;
		}

		commitParseToTree(editedXml, false);
	}, [editedXml, onParseToTree, report, commitParseToTree]);

	const handleConfirmParseToTree = useCallback(() => {
		if (!pendingParse) return;
		const { xml: xmlToParse } = pendingParse;
		setPendingParse(null);
		commitParseToTree(xmlToParse, true);
	}, [pendingParse, commitParseToTree]);

	// Handle editor mount
	const handleEditorMount = useCallback(
		(editor: MonacoEditor.editor.IStandaloneCodeEditor, monaco: typeof MonacoEditor) => {
			editorRef.current = editor;
			// FetchXML completions and hovers for this editor only, released with it (#194)
			const languageFeatures = attachFetchXmlLanguageFeatures(monaco, editor);
			editor.onDidDispose(() => languageFeatures.dispose());
		},
		[],
	);

	// Bar order; what overflows first: Load 10, Reset 15, Paste 20, Format 30, Parse to tree 35,
	// Validate 40, Copy 50 (#232, #186). Tooltips describe; the label names.
	const copyCommand: BarCommand = {
		id: "copy",
		label: "Copy",
		icon: CopyIcon,
		group: "clipboard",
		priority: 50,
		kind: "button",
		hint: "Copy FetchXML to clipboard",
		onClick: handleCopy,
	};
	const editorCommands: BarCommand[] = isEditorMode
		? [
				copyCommand,
				{
					id: "paste",
					label: "Paste",
					icon: PasteIcon,
					group: "clipboard",
					priority: 20,
					kind: "button",
					hint: "Paste from clipboard",
					onClick: handlePaste,
				},
				{
					id: "load",
					label: "Load",
					icon: LoadIcon,
					group: "clipboard",
					priority: 10,
					kind: "button",
					hint: "Load from file",
					onClick: handleLoadFile,
				},
				{
					id: "reset",
					label: "Reset",
					icon: ResetIcon,
					group: "clipboard",
					priority: 15,
					kind: "button",
					hint: "Reset to tree-generated XML",
					onClick: handleReset,
				},
				{
					id: "format",
					label: "Format",
					icon: FormatIcon,
					group: "check",
					priority: 30,
					kind: "button",
					hint: `Re-indent the FetchXML (${formatKeys(getEntry("formatXml"))[0]})`,
					keyShortcuts: getEntry("formatXml").aria,
					disabledReason: editedXml.trim() ? undefined : "There's no FetchXML to format",
					onClick: handleFormat,
				},
				{
					id: "validate",
					label: "Validate",
					icon: ValidateIcon,
					group: "check",
					priority: 40,
					kind: "button",
					hint: "Validate FetchXML syntax",
					onClick: handleValidate,
				},
				{
					id: "parse",
					label: "Parse to tree",
					icon: ParseIcon,
					group: "check",
					priority: 35,
					kind: "button",
					hint: "Parse FetchXML into tree builder",
					disabledReason: onParseToTree ? undefined : "The query tree isn't available",
					onClick: handleParseToTree,
				},
			]
		: [copyCommand];

	return (
		<div className={styles.container}>
			{/* The editor's commands (#232): Copy only in view mode. The Switch sits beside the bar and
			    never overflows. */}
			<div className={styles.toolbar}>
				<OverflowToolbar
					ariaLabel="FetchXML commands"
					size="small"
					groups={["clipboard", "check"]}
					className={styles.commands}
					commands={editorCommands}
				/>
				<Switch
					checked={isEditorMode}
					onChange={(_, data) => handleToggleEditorMode(data.checked)}
					label={<span className={styles.switchLabel}>Editor mode</span>}
					labelPosition="before"
					className={styles.modeSwitch}
				/>
			</div>

			{/* The editor's bars, one per condition, under the toolbar (#277) */}
			{messages.length > 0 && (
				<div className={styles.messageBarContainer}>
					<MessageBarGroup animate="both">
						{messages.map((msg) => (
							<MessageBar
								key={msg.key}
								intent={msg.intent}
								layout={typeof msg.body === "string" ? "auto" : "multiline"}
							>
								<MessageBarBody>
									<MessageBarTitle>{msg.title}</MessageBarTitle>
									{typeof msg.body === "string" ? (
										msg.body
									) : (
										<ParseIssueList items={msg.body} />
									)}
								</MessageBarBody>
								<MessageBarActions
									containerAction={
										<Button
											onClick={() => dismissMessage(msg.key)}
											aria-label="Dismiss"
											appearance="transparent"
											icon={<Dismiss16Regular />}
										/>
									}
								>
									{msg.action && (
										<Button onClick={msg.action.onClick}>
											{msg.action.label}
										</Button>
									)}
									{msg.advisories && onViewAdvisories && (
										<Button onClick={onViewAdvisories}>
											View all in Advisories
										</Button>
									)}
								</MessageBarActions>
							</MessageBar>
						))}
					</MessageBarGroup>
				</div>
			)}

			{/* Hidden file input */}
			<input
				ref={fileInputRef}
				type="file"
				accept=".xml,.fetchxml,.txt"
				style={{ display: "none" }}
				onChange={handleFileChange}
			/>

			{/* Editor */}
			<div className={styles.editorWrapper}>
				<Editor
					height="100%"
					language="xml"
					theme={isDark ? "vs-dark" : "vs"}
					value={currentXml}
					onChange={
						isEditorMode
							? (value) => {
									setEditedXml(value || "");
									onEditorStateChange?.(true, value || "");
								}
							: undefined
					}
					loading={
						<div className={styles.loadingContainer}>
							<Spinner size="medium" label="Loading editor…" />
						</div>
					}
					options={{
						readOnly: !isEditorMode,
						minimap: { enabled: false },
						scrollBeyondLastLine: false,
						lineNumbers: "on",
						folding: true,
						wordWrap: "on",
						automaticLayout: true,
						fontFamily: "Consolas, 'Courier New', monospace",
						fontSize: 13,
					}}
					onMount={handleEditorMount}
				/>
			</div>

			{/* Parse to tree confirmation (#89) */}
			<ParseChangesDialog
				action={pendingParse ? "parse" : null}
				changes={pendingParse?.changes ?? []}
				onConfirm={handleConfirmParseToTree}
				onCancel={() => setPendingParse(null)}
			/>
		</div>
	);
}

export const FetchXmlEditor = forwardRef(FetchXmlEditorImpl);
