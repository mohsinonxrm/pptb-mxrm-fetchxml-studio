/**
 * Monaco editor for FetchXML with Editor Mode toggle
 * - Read-only mode: Shows generated XML with Copy button
 * - Editor mode: Editable with Paste, Load File, Reset, Validate, Parse to Tree buttons
 */

import { useState, useRef, useCallback, useEffect } from "react";
import Editor from "@monaco-editor/react";
import type * as MonacoEditor from "monaco-editor";
import {
	Button,
	makeStyles,
	Tooltip,
	Spinner,
	tokens,
	Switch,
	Toolbar,
	ToolbarButton,
	ToolbarDivider,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	MessageBarActions,
	MessageBarGroup,
} from "@fluentui/react-components";
import {
	Copy24Regular,
	ClipboardPaste24Regular,
	FolderOpen24Regular,
	ArrowReset24Regular,
	Checkmark24Regular,
	TreeDeciduous24Regular,
	Dismiss24Regular,
} from "@fluentui/react-icons";
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
import { ParseIssueList, type ParseIssueItem } from "./ParseIssueList";
import { ParseChangesDialog } from "./ParseChangesDialog";
import {
	registerFetchXmlIntellisense,
	registerFetchXmlHoverProvider,
} from "../../model/fetchxmlIntellisense";

/** FetchXML as the tree would write it, to compare meaning rather than formatting; null if it doesn't parse. */
function canonicalFetchXml(xmlText: string): string | null {
	const result = parseFetchXml(xmlText);
	return result.success && result.fetchNode ? generateFetchXml(result.fetchNode) : null;
}

const TREE_CHANGED_TITLE = "The tree changed while you were editing";

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
		justifyContent: "space-between",
		padding: "4px 8px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
		backgroundColor: tokens.colorNeutralBackground2,
		flexShrink: 0,
	},
	toolbarLeft: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
	},
	toolbarRight: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
	},
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
		borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
		maxHeight: "150px",
		overflowY: "auto",
		flexShrink: 0,
	},
	switchLabel: {
		fontSize: tokens.fontSizeBase200,
		fontWeight: tokens.fontWeightSemibold,
	},
});

interface Message {
	id: number;
	intent: "success" | "warning" | "error" | "info";
	title: string;
	/** Text, or a list of findings rendered as bullets */
	body: string | ParseIssueItem[];
}

const INVALID_TITLE = "Invalid FetchXML";
const MIGHT_REJECT_TITLE = "Dataverse might reject this query";

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
}

export function FetchXmlEditor({
	xml,
	onParseToTree,
	onEditorStateChange,
	replaceRequest,
}: FetchXmlEditorProps) {
	const styles = useStyles();
	const { isDark } = useTheme();
	const editorRef = useRef<MonacoEditor.editor.IStandaloneCodeEditor | null>(null);
	const fileInputRef = useRef<HTMLInputElement>(null);
	const messageIdRef = useRef(0);

	// Editor mode state
	const [isEditorMode, setIsEditorMode] = useState(false);
	const [editedXml, setEditedXml] = useState("");
	const [messages, setMessages] = useState<Message[]>([]);
	// Parse to Tree awaiting confirmation because it would change the query (#89)
	const [pendingParse, setPendingParse] = useState<{
		xml: string;
		changes: ParseWarning[];
	} | null>(null);

	// Get current XML value based on mode
	const currentXml = isEditorMode ? editedXml : xml;

	// Add a message
	const addMessage = useCallback(
		(intent: Message["intent"], title: string, body: Message["body"]) => {
		const newMessage: Message = {
			id: messageIdRef.current++,
			intent,
			title,
			body,
		};
			setMessages((prev) => [newMessage, ...prev].slice(0, 5)); // Keep max 5 messages
		},
		[],
	);

	// Dismiss a message
	const dismissMessage = useCallback((id: number) => {
		setMessages((prev) => prev.filter((m) => m.id !== id));
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
		setMessages((prev) =>
			prev.some((m) => m.title === TREE_CHANGED_TITLE)
				? prev
				: [
						{
							id: messageIdRef.current++,
							intent: "warning" as const,
							title: TREE_CHANGED_TITLE,
							body: "Execute, Save and Send to Tool use the editor's FetchXML and will replace the tree's changes. Use Reset to take the tree's FetchXML instead.",
						},
						...prev,
					].slice(0, 5),
		);
		// Only tree changes drive this; the editor's own edits are compared when the tree changes
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [xml, isEditorMode]);

	// Show what the tree loaded when an editor-mode action changed the query (#102)
	useEffect(() => {
		if (!replaceRequest || !isEditorMode) return;
		setEditedXml(replaceRequest.xml);
		onEditorStateChange?.(true, replaceRequest.xml);
		addMessage(
			"info",
			"Editor updated",
			"The editor now shows the FetchXML that was loaded into the tree, with the changes you confirmed.",
		);
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
				onEditorStateChange?.(false, "");
			}
			setIsEditorMode(checked);
		},
		[xml, clearMessages, onEditorStateChange],
	);

	// Copy to clipboard
	const handleCopy = useCallback(async () => {
		try {
			await navigator.clipboard.writeText(currentXml);
			addMessage("success", "Copied", "FetchXML copied to clipboard");
		} catch (err) {
			console.error("Failed to copy:", err);
			addMessage("error", "Copy failed", "Could not copy to clipboard");
		}
	}, [currentXml, addMessage]);

	// Paste from clipboard
	const handlePaste = useCallback(async () => {
		try {
			const text = await navigator.clipboard.readText();
			if (text.trim()) {
				setEditedXml(text);
				onEditorStateChange?.(true, text);
				addMessage("info", "Pasted", "Content pasted from clipboard");
			}
		} catch (err) {
			console.error("Failed to paste:", err);
			addMessage("error", "Paste failed", "Could not read from clipboard");
		}
	}, [addMessage, onEditorStateChange]);

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
					addMessage("info", "File loaded", `Loaded ${file.name}`);
				};
				reader.onerror = () => {
					addMessage("error", "Load failed", "Could not read file");
				};
				reader.readAsText(file);
			}
			// Reset input so same file can be loaded again
			event.target.value = "";
		},
		[addMessage, onEditorStateChange],
	);

	// Reset to tree-generated XML
	const handleReset = useCallback(() => {
		setEditedXml(xml);
		onEditorStateChange?.(true, xml);
		clearMessages();
		addMessage("info", "Reset", "Reverted to tree-generated FetchXML");
	}, [xml, clearMessages, addMessage, onEditorStateChange]);

	// Validate XML
	const handleValidate = useCallback(() => {
		clearMessages();

		// First check syntax
		const syntaxResult = validateFetchXmlSyntax(editedXml);
		if (!syntaxResult.valid) {
			addMessage("error", "Syntax Error", syntaxResult.error || "Invalid XML syntax");
			return;
		}

		// Then do a full parse to classify findings
		const parseResult = parseFetchXml(editedXml);
		if (!parseResult.success) {
			addMessage("error", INVALID_TITLE, parseResult.errors);
			return;
		}

		// One bar per level. Messages are prepended, so add the least severe first.
		const notes = parseResult.warnings.filter((w) => w.severity === "info");
		const blocking = getBlockingWarnings(parseResult);
		const executionErrors = getExecutionErrors(parseResult);
		if (notes.length > 0) addMessage("info", "Notes", notes);
		if (blocking.length > 0) addMessage("warning", MIGHT_REJECT_TITLE, blocking);
		if (executionErrors.length > 0) addMessage("error", INVALID_TITLE, executionErrors);
		if (blocking.length === 0 && executionErrors.length === 0) {
			addMessage("success", "Valid", "Dataverse should accept this FetchXML.");
		}
	}, [editedXml, clearMessages, addMessage]);

	// Load XML into the tree. `changesConfirmed` is true when the user already reviewed the
	// changes in the confirmation dialog, so they aren't repeated as notes.
	const commitParseToTree = useCallback(
		(xmlToParse: string, changesConfirmed: boolean) => {
			if (!onParseToTree) return;

			clearMessages();
			const result = onParseToTree(xmlToParse);

			if (!result.success) {
				addMessage("error", INVALID_TITLE, result.errors);
				return;
			}

			const changes = new Set(getParseToTreeChanges(result));
			// Kept as written by Parse to Tree, so still a problem in the tree
			const stillInvalid = result.warnings.filter((w) => w.kind === "invalid");
			const stillErrors = stillInvalid.filter((w) => w.severity === "error");
			const stillWarnings = stillInvalid.filter((w) => w.severity === "warning");
			const notes = result.warnings.filter(
				(w) => w.severity === "info" && !(changesConfirmed && changes.has(w)),
			);
			if (notes.length > 0) addMessage("info", "Notes", notes);
			if (stillWarnings.length > 0) addMessage("warning", MIGHT_REJECT_TITLE, stillWarnings);
			if (stillErrors.length > 0) addMessage("error", INVALID_TITLE, stillErrors);
			addMessage("success", "Parsed into the tree", "The FetchXML was loaded into the tree builder.");

			// Exit editor mode after successful parse
			onEditorStateChange?.(false, "");
			setIsEditorMode(false);
		},
		[onParseToTree, clearMessages, addMessage, onEditorStateChange],
	);

	// Parse to tree — confirm first when the tree would change the query
	const handleParseToTree = useCallback(() => {
		if (!onParseToTree) return;

		const preview = parseFetchXml(editedXml);
		if (!preview.success) {
			clearMessages();
			addMessage("error", INVALID_TITLE, preview.errors);
			return;
		}

		const changes = getParseToTreeChanges(preview);
		if (changes.length > 0) {
			setPendingParse({ xml: editedXml, changes });
			return;
		}

		commitParseToTree(editedXml, false);
	}, [editedXml, onParseToTree, clearMessages, addMessage, commitParseToTree]);

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
			// Register FetchXML intellisense (only in editor mode, but register once)
			registerFetchXmlIntellisense(monaco);
			registerFetchXmlHoverProvider(monaco);
		},
		[],
	);

	return (
		<div className={styles.container}>
			{/* Toolbar */}
			<div className={styles.toolbar}>
				<div className={styles.toolbarLeft}>
					{isEditorMode ? (
						<Toolbar size="small">
							<Tooltip content="Copy FetchXML to clipboard" relationship="label">
								<ToolbarButton icon={<Copy24Regular />} onClick={handleCopy}>
									Copy
								</ToolbarButton>
							</Tooltip>
							<Tooltip content="Paste from clipboard" relationship="label">
								<ToolbarButton icon={<ClipboardPaste24Regular />} onClick={handlePaste}>
									Paste
								</ToolbarButton>
							</Tooltip>
							<Tooltip content="Load from file" relationship="label">
								<ToolbarButton icon={<FolderOpen24Regular />} onClick={handleLoadFile}>
									Load
								</ToolbarButton>
							</Tooltip>
							<Tooltip content="Reset to tree-generated XML" relationship="label">
								<ToolbarButton icon={<ArrowReset24Regular />} onClick={handleReset}>
									Reset
								</ToolbarButton>
							</Tooltip>
							<ToolbarDivider />
							<Tooltip content="Validate FetchXML syntax" relationship="label">
								<ToolbarButton icon={<Checkmark24Regular />} onClick={handleValidate}>
									Validate
								</ToolbarButton>
							</Tooltip>
							<Tooltip content="Parse FetchXML into tree builder" relationship="label">
								<ToolbarButton
									icon={<TreeDeciduous24Regular />}
									onClick={handleParseToTree}
									disabled={!onParseToTree}
								>
									Parse to Tree
								</ToolbarButton>
							</Tooltip>
						</Toolbar>
					) : (
						<Tooltip content="Copy FetchXML to clipboard" relationship="label">
							<Button appearance="subtle" icon={<Copy24Regular />} onClick={handleCopy}>
								Copy
							</Button>
						</Tooltip>
					)}
				</div>
				<div className={styles.toolbarRight}>
					<Switch
						checked={isEditorMode}
						onChange={(_, data) => handleToggleEditorMode(data.checked)}
						label={<span className={styles.switchLabel}>Editor Mode</span>}
						labelPosition="before"
					/>
				</div>
			</div>

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
							<Spinner size="medium" label="Loading editor..." />
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

			{/* Messages */}
			{messages.length > 0 && (
				<div className={styles.messageBarContainer}>
					<MessageBarGroup animate="both">
						{messages.map((msg) => (
							<MessageBar
								key={msg.id}
								intent={msg.intent}
								shape="rounded"
								layout={typeof msg.body === "string" ? "auto" : "multiline"}
							>
								<MessageBarBody>
									<MessageBarTitle>{msg.title}</MessageBarTitle>
									{typeof msg.body === "string" ? msg.body : <ParseIssueList items={msg.body} />}
								</MessageBarBody>
								<MessageBarActions
									containerAction={
										<Button
											onClick={() => dismissMessage(msg.id)}
											aria-label="dismiss"
											appearance="transparent"
											icon={<Dismiss24Regular />}
											size="small"
										/>
									}
								/>
							</MessageBar>
						))}
					</MessageBarGroup>
				</div>
			)}

			{/* Parse to Tree confirmation (#89) */}
			<ParseChangesDialog
				action={pendingParse ? "parse" : null}
				changes={pendingParse?.changes ?? []}
				onConfirm={handleConfirmParseToTree}
				onCancel={() => setPendingParse(null)}
			/>
		</div>
	);
}
