/**
 * Monaco editor for FetchXML preview with syntax highlighting and copy functionality
 */

import { useEffect, useRef } from "react";
import Editor from "@monaco-editor/react";
import { Button, makeStyles, Tooltip } from "@fluentui/react-components";
import { Copy24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		position: "relative",
	},
	editorWrapper: {
		flex: 1,
		overflow: "hidden",
	},
	copyButton: {
		position: "absolute",
		top: "8px",
		right: "8px",
		zIndex: 10,
	},
});

interface FetchXmlEditorProps {
	xml: string;
	isDark: boolean;
}

export function FetchXmlEditor({ xml, isDark }: FetchXmlEditorProps) {
	const styles = useStyles();
	const editorRef = useRef<{ getAction: (id: string) => { run: () => void } | null } | null>(null);

	const handleCopy = async () => {
		try {
			await navigator.clipboard.writeText(xml);
			// Could show a toast notification here via PPTB API
		} catch (err) {
			console.error("Failed to copy:", err);
		}
	};

	useEffect(() => {
		// Format XML when editor is ready
		if (editorRef.current) {
			editorRef.current.getAction("editor.action.formatDocument")?.run();
		}
	}, [xml]);

	return (
		<div className={styles.container}>
			<Tooltip content="Copy FetchXML to clipboard" relationship="label">
				<Button
					className={styles.copyButton}
					appearance="subtle"
					icon={<Copy24Regular />}
					onClick={handleCopy}
				/>
			</Tooltip>
			<div className={styles.editorWrapper}>
				<Editor
					height="100%"
					language="xml"
					theme={isDark ? "vs-dark" : "vs"}
					value={xml}
					options={{
						readOnly: true,
						minimap: { enabled: false },
						scrollBeyondLastLine: false,
						lineNumbers: "on",
						folding: true,
						wordWrap: "on",
						automaticLayout: true,
					}}
					onMount={(editor) => {
						editorRef.current = editor;
						// Auto-format on mount
						setTimeout(() => {
							editor.getAction("editor.action.formatDocument")?.run();
						}, 100);
					}}
				/>
			</div>
		</div>
	);
}
