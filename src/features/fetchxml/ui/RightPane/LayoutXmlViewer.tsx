/**
 * Monaco editor for LayoutXML (read-only)
 * Displays the LayoutXML that defines column configuration for the grid
 */

import { useRef, useCallback } from "react";
import Editor from "@monaco-editor/react";
import type * as MonacoEditor from "monaco-editor/editor";
import { makeStyles, Spinner, tokens } from "@fluentui/react-components";
import { Copy20Filled, Copy20Regular, bundleIcon } from "@fluentui/react-icons";
import { OverflowToolbar } from "../../../../shared/components/commandBar";
import { useCopy } from "../../../../shared/feedback/useCopy";
import { useTheme } from "../../../../shared/contexts/useTheme";

const CopyIcon = bundleIcon(Copy20Filled, Copy20Regular);

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		overflow: "hidden",
	},
	toolbar: {
		paddingInline: tokens.spacingHorizontalS,
		paddingBlock: tokens.spacingVerticalXXS,
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
		backgroundColor: tokens.colorNeutralBackground2,
		flexShrink: 0,
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

	// "Copied LayoutXML" or "Couldn't copy to the clipboard" (#277)
	const copy = useCopy();
	const handleCopy = useCallback(async () => {
		await copy(formattedXml, "LayoutXML");
	}, [formattedXml, copy]);

	return (
		<div className={styles.container}>
			{/* The label, then Copy at the end (#232) */}
			<OverflowToolbar
				ariaLabel="LayoutXML commands"
				size="small"
				align="end"
				className={styles.toolbar}
				start={<span className={styles.label}>LayoutXML (read-only)</span>}
				commands={[
					{
						id: "copy",
						label: "Copy",
						ariaLabel: "Copy LayoutXML",
						icon: CopyIcon,
						kind: "button",
						priority: 10,
						disabledReason: layoutXml ? undefined : "There's no LayoutXML to copy yet",
						onClick: () => void handleCopy(),
					},
				]}
			/>
			<div className={styles.editorWrapper}>
				<Editor
					height="100%"
					language="xml"
					theme={isDark ? "vs-dark" : "vs"}
					value={formattedXml}
					options={{
						ariaLabel: "LayoutXML, read-only",
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
							<Spinner size="medium" label="Loading editor…" />
						</div>
					}
				/>
			</div>
		</div>
	);
}
