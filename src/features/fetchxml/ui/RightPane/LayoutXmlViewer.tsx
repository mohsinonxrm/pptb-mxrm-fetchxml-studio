/**
 * Monaco editor for LayoutXML (read-only)
 * Displays the LayoutXML that defines column configuration for the grid
 */

import { useRef, useCallback } from "react";
import Editor from "@monaco-editor/react";
import type * as MonacoEditor from "monaco-editor";
import { makeStyles, Spinner, tokens, Toolbar, ToolbarButton } from "@fluentui/react-components";
import { Copy24Regular } from "@fluentui/react-icons";
import { useTheme } from "../../../../shared/contexts/ThemeContext";

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
	label: {
		fontSize: tokens.fontSizeBase200,
		fontWeight: tokens.fontWeightSemibold,
		color: tokens.colorNeutralForeground2,
	},
});

interface LayoutXmlViewerProps {
	/** The LayoutXML string to display */
	layoutXml: string;
}

export function LayoutXmlViewer({ layoutXml }: LayoutXmlViewerProps) {
	const styles = useStyles();
	const { isDark } = useTheme();
	const editorRef = useRef<MonacoEditor.editor.IStandaloneCodeEditor | null>(null);

	// Format XML with indentation for better readability
	const formatXml = (xml: string): string => {
		if (!xml) return "";
		try {
			// Simple XML formatting - add newlines and indentation
			let formatted = "";
			let indent = 0;
			const parts = xml.replace(/>\s*</g, ">\n<").split("\n");

			for (const part of parts) {
				const trimmed = part.trim();
				if (!trimmed) continue;

				// Decrease indent for closing tags
				if (trimmed.startsWith("</")) {
					indent = Math.max(0, indent - 1);
				}

				formatted += "  ".repeat(indent) + trimmed + "\n";

				// Increase indent for opening tags (not self-closing)
				if (
					trimmed.startsWith("<") &&
					!trimmed.startsWith("</") &&
					!trimmed.startsWith("<?") &&
					!trimmed.endsWith("/>") &&
					!trimmed.includes("</")
				) {
					indent++;
				}
			}

			return formatted.trim();
		} catch {
			return xml;
		}
	};

	const formattedXml = formatXml(layoutXml);

	const handleEditorMount = useCallback((editor: MonacoEditor.editor.IStandaloneCodeEditor) => {
		editorRef.current = editor;
	}, []);

	const handleCopy = useCallback(async () => {
		try {
			await navigator.clipboard.writeText(formattedXml);
		} catch (error) {
			console.error("Failed to copy to clipboard:", error);
		}
	}, [formattedXml]);

	return (
		<div className={styles.container}>
			<div className={styles.toolbar}>
				<div className={styles.toolbarLeft}>
					<span className={styles.label}>LayoutXML (Read-only)</span>
				</div>
				<Toolbar size="small">
					<ToolbarButton
						icon={<Copy24Regular />}
						onClick={handleCopy}
						disabled={!layoutXml}
						aria-label="Copy LayoutXML"
					>
						Copy
					</ToolbarButton>
				</Toolbar>
			</div>
			<div className={styles.editorWrapper}>
				<Editor
					height="100%"
					language="xml"
					theme={isDark ? "vs-dark" : "vs"}
					value={formattedXml}
					options={{
						readOnly: true,
						minimap: { enabled: false },
						lineNumbers: "on",
						scrollBeyondLastLine: false,
						wordWrap: "on",
						fontSize: 13,
						tabSize: 2,
						automaticLayout: true,
						folding: true,
						renderLineHighlight: "line",
						scrollbar: {
							vertical: "visible",
							horizontal: "visible",
							verticalScrollbarSize: 10,
							horizontalScrollbarSize: 10,
						},
					}}
					onMount={handleEditorMount}
					loading={
						<div className={styles.loadingContainer}>
							<Spinner size="medium" label="Loading editor..." />
						</div>
					}
				/>
			</div>
		</div>
	);
}
