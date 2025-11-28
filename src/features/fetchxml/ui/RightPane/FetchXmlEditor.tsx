/**
 * Monaco editor for FetchXML with Editor Mode toggle
 * - Read-only mode: Shows generated XML with Copy button
 * - Editor mode: Editable with Paste, Load File, Reset, Validate, Parse to Tree buttons
 */

import { useState, useRef, useCallback } from "react";
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
import { useTheme } from "../../../../shared/contexts/ThemeContext";
import { parseFetchXml, validateFetchXmlSyntax } from "../../model/fetchxmlParser";
import type { ParseResult, ParseWarning, ParseError } from "../../model/fetchxmlParser";
import {
	registerFetchXmlIntellisense,
	registerFetchXmlHoverProvider,
} from "../../model/fetchxmlIntellisense";

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
	body: string;
}

interface FetchXmlEditorProps {
	xml: string;
	onParseToTree?: (xmlString: string) => ParseResult;
}

export function FetchXmlEditor({ xml, onParseToTree }: FetchXmlEditorProps) {
	const styles = useStyles();
	const { isDark } = useTheme();
	const editorRef = useRef<MonacoEditor.editor.IStandaloneCodeEditor | null>(null);
	const fileInputRef = useRef<HTMLInputElement>(null);
	const messageIdRef = useRef(0);

	// Editor mode state
	const [isEditorMode, setIsEditorMode] = useState(false);
	const [editedXml, setEditedXml] = useState("");
	const [messages, setMessages] = useState<Message[]>([]);

	// Get current XML value based on mode
	const currentXml = isEditorMode ? editedXml : xml;

	// Add a message
	const addMessage = useCallback((intent: Message["intent"], title: string, body: string) => {
		const newMessage: Message = {
			id: messageIdRef.current++,
			intent,
			title,
			body,
		};
		setMessages((prev) => [newMessage, ...prev].slice(0, 5)); // Keep max 5 messages
	}, []);

	// Dismiss a message
	const dismissMessage = useCallback((id: number) => {
		setMessages((prev) => prev.filter((m) => m.id !== id));
	}, []);

	// Clear all messages
	const clearMessages = useCallback(() => {
		setMessages([]);
	}, []);

	// Toggle editor mode
	const handleToggleEditorMode = useCallback(
		(checked: boolean) => {
			if (checked) {
				// Entering editor mode - copy current XML to edit buffer
				setEditedXml(xml);
				clearMessages();
			}
			setIsEditorMode(checked);
		},
		[xml, clearMessages]
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
				addMessage("info", "Pasted", "Content pasted from clipboard");
			}
		} catch (err) {
			console.error("Failed to paste:", err);
			addMessage("error", "Paste failed", "Could not read from clipboard");
		}
	}, [addMessage]);

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
		[addMessage]
	);

	// Reset to tree-generated XML
	const handleReset = useCallback(() => {
		setEditedXml(xml);
		clearMessages();
		addMessage("info", "Reset", "Reverted to tree-generated FetchXML");
	}, [xml, clearMessages, addMessage]);

	// Validate XML
	const handleValidate = useCallback(() => {
		clearMessages();

		// First check syntax
		const syntaxResult = validateFetchXmlSyntax(editedXml);
		if (!syntaxResult.valid) {
			addMessage("error", "Syntax Error", syntaxResult.error || "Invalid XML syntax");
			return;
		}

		// Then do a full parse to get warnings
		const parseResult = parseFetchXml(editedXml);
		if (!parseResult.success) {
			parseResult.errors.forEach((err: ParseError) => {
				addMessage("error", "Parse Error", err.message);
			});
			return;
		}

		// Show warnings if any
		if (parseResult.warnings.length > 0) {
			parseResult.warnings.forEach((warn: ParseWarning) => {
				const prefix = warn.element ? `<${warn.element}>: ` : "";
				addMessage("warning", "Warning", `${prefix}${warn.message}`);
			});
		}

		addMessage("success", "Valid", "FetchXML is valid and ready to parse");
	}, [editedXml, clearMessages, addMessage]);

	// Parse to tree
	const handleParseToTree = useCallback(() => {
		if (!onParseToTree) return;

		clearMessages();
		const result = onParseToTree(editedXml);

		if (result.success) {
			// Show warnings but still success
			result.warnings.forEach((warn: ParseWarning) => {
				const prefix = warn.element ? `<${warn.element}>: ` : "";
				addMessage("warning", "Warning", `${prefix}${warn.message}`);
			});
			addMessage("success", "Parsed", "FetchXML loaded into tree builder");
			// Exit editor mode after successful parse
			setIsEditorMode(false);
		} else {
			result.errors.forEach((err: ParseError) => {
				addMessage("error", "Parse Error", err.message);
			});
		}
	}, [editedXml, onParseToTree, clearMessages, addMessage]);

	// Handle editor mount
	const handleEditorMount = useCallback(
		(editor: MonacoEditor.editor.IStandaloneCodeEditor, monaco: typeof MonacoEditor) => {
			editorRef.current = editor;
			// Register FetchXML intellisense (only in editor mode, but register once)
			registerFetchXmlIntellisense(monaco);
			registerFetchXmlHoverProvider(monaco);
		},
		[]
	);

	return (
		<div className={styles.container}>
			{/* Toolbar */}
			<div className={styles.toolbar}>
				<div className={styles.toolbarLeft}>
					{isEditorMode ? (
						<Toolbar size="small">
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
					onChange={isEditorMode ? (value) => setEditedXml(value || "") : undefined}
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
							<MessageBar key={msg.id} intent={msg.intent} shape="rounded">
								<MessageBarBody>
									<MessageBarTitle>{msg.title}</MessageBarTitle>
									{msg.body}
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
		</div>
	);
}
