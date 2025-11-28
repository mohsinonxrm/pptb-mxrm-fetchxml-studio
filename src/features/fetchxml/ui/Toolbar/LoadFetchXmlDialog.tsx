/**
 * Dialog for loading FetchXML from text input
 * Uses Monaco editor with FetchXML intellisense
 */

import { useState, useCallback, useRef } from "react";
import {
	Dialog,
	DialogSurface,
	DialogBody,
	DialogTitle,
	DialogContent,
	DialogActions,
	Button,
	makeStyles,
	tokens,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Spinner,
} from "@fluentui/react-components";
import { Warning20Regular, ErrorCircle20Regular } from "@fluentui/react-icons";
import Editor from "@monaco-editor/react";
import type * as MonacoEditor from "monaco-editor";
import { parseFetchXml, validateFetchXmlSyntax } from "../../model/fetchxmlParser";
import type { ParseResult, ParseWarning } from "../../model/fetchxmlParser";
import {
	registerFetchXmlIntellisense,
	registerFetchXmlHoverProvider,
} from "../../model/fetchxmlIntellisense";
import { useTheme } from "../../../../shared/contexts/ThemeContext";

const useStyles = makeStyles({
	content: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
	},
	textarea: {
		minHeight: "300px",
		fontFamily: "Consolas, 'Courier New', monospace",
		fontSize: "13px",
	},
	warningList: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
		maxHeight: "120px",
		overflowY: "auto",
	},
	warningItem: {
		fontSize: "12px",
		color: tokens.colorPaletteMarigoldForeground1,
		display: "flex",
		alignItems: "flex-start",
		gap: tokens.spacingHorizontalXS,
	},
	errorItem: {
		fontSize: "12px",
		color: tokens.colorPaletteRedForeground1,
		display: "flex",
		alignItems: "flex-start",
		gap: tokens.spacingHorizontalXS,
	},
	helpText: {
		fontSize: "12px",
		color: tokens.colorNeutralForeground3,
	},
	loadingContainer: {
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		height: "300px",
		backgroundColor: tokens.colorNeutralBackground1,
	},
});

interface LoadFetchXmlDialogProps {
	open: boolean;
	onClose: () => void;
	onLoad: (xmlString: string) => ParseResult;
}

export function LoadFetchXmlDialog({ open, onClose, onLoad }: LoadFetchXmlDialogProps) {
	const styles = useStyles();
	const { isDark } = useTheme();

	const [xmlInput, setXmlInput] = useState("");
	const [parseErrors, setParseErrors] = useState<string[]>([]);
	const [parseWarnings, setParseWarnings] = useState<ParseWarning[]>([]);
	const [isValidating, setIsValidating] = useState(false);
	const editorRef = useRef<MonacoEditor.editor.IStandaloneCodeEditor | null>(null);

	const handleClose = useCallback(() => {
		setXmlInput("");
		setParseErrors([]);
		setParseWarnings([]);
		onClose();
	}, [onClose]);

	const handleValidate = useCallback(() => {
		setIsValidating(true);
		setParseErrors([]);
		setParseWarnings([]);

		// Quick syntax validation
		const syntaxResult = validateFetchXmlSyntax(xmlInput);
		if (!syntaxResult.valid) {
			setParseErrors([syntaxResult.error || "Invalid FetchXML syntax"]);
			setIsValidating(false);
			return;
		}

		// Full parse to check for warnings
		const parseResult = parseFetchXml(xmlInput);
		if (!parseResult.success) {
			setParseErrors(parseResult.errors.map((e) => e.message));
		} else {
			setParseWarnings(parseResult.warnings);
		}

		setIsValidating(false);
	}, [xmlInput]);

	const handleLoad = useCallback(() => {
		const result = onLoad(xmlInput);

		if (result.success) {
			// Show warnings if any but still close
			if (result.warnings.length > 0) {
				console.log("FetchXML loaded with warnings:", result.warnings);
			}
			handleClose();
		} else {
			// Show errors and don't close
			setParseErrors(result.errors.map((e) => e.message));
		}
	}, [xmlInput, onLoad, handleClose]);

	const hasErrors = parseErrors.length > 0;
	const hasWarnings = parseWarnings.length > 0;
	const hasInput = xmlInput.trim().length > 0;

	return (
		<Dialog open={open} onOpenChange={(_, data) => !data.open && handleClose()}>
			<DialogSurface style={{ maxWidth: "700px", width: "90vw" }}>
				<DialogBody>
					<DialogTitle>Load FetchXML</DialogTitle>
					<DialogContent className={styles.content}>
						<p className={styles.helpText}>
							Paste your FetchXML below. It will be parsed and loaded into the builder tree.
						</p>

						<div
							style={{
								border: `1px solid ${tokens.colorNeutralStroke1}`,
								borderRadius: tokens.borderRadiusMedium,
							}}
						>
							<Editor
								height="300px"
								language="xml"
								theme={isDark ? "vs-dark" : "vs"}
								value={xmlInput}
								loading={
									<div className={styles.loadingContainer}>
										<Spinner size="medium" label="Loading editor..." />
									</div>
								}
								onChange={(value) => {
									setXmlInput(value || "");
									// Clear validation on input change
									setParseErrors([]);
									setParseWarnings([]);
								}}
								onMount={(editor, monaco) => {
									// Register FetchXML intellisense
									registerFetchXmlIntellisense(monaco);
									registerFetchXmlHoverProvider(monaco);
									editorRef.current = editor;
								}}
								options={{
									minimap: { enabled: false },
									lineNumbers: "on",
									folding: true,
									wordWrap: "on",
									automaticLayout: true,
									scrollBeyondLastLine: false,
									fontFamily: "Consolas, 'Courier New', monospace",
									fontSize: 13,
									placeholder:
										'<fetch>\n  <entity name="account">\n    <attribute name="name" />\n  </entity>\n</fetch>',
								}}
							/>
						</div>

						{/* Errors */}
						{hasErrors && (
							<MessageBar intent="error">
								<MessageBarBody>
									<MessageBarTitle>Parsing Errors</MessageBarTitle>
									<div className={styles.warningList}>
										{parseErrors.map((error, i) => (
											<div key={i} className={styles.errorItem}>
												<ErrorCircle20Regular />
												<span>{error}</span>
											</div>
										))}
									</div>
								</MessageBarBody>
							</MessageBar>
						)}

						{/* Warnings */}
						{hasWarnings && !hasErrors && (
							<MessageBar intent="warning">
								<MessageBarBody>
									<MessageBarTitle>Validation Warnings</MessageBarTitle>
									<div className={styles.warningList}>
										{parseWarnings.map((warning, i) => (
											<div key={i} className={styles.warningItem}>
												<Warning20Regular />
												<span>
													{warning.element ? `<${warning.element}>: ` : ""}
													{warning.message}
												</span>
											</div>
										))}
									</div>
								</MessageBarBody>
							</MessageBar>
						)}
					</DialogContent>
					<DialogActions>
						<Button appearance="secondary" onClick={handleValidate} disabled={!hasInput}>
							Validate
						</Button>
						<Button
							appearance="primary"
							onClick={handleLoad}
							disabled={!hasInput || hasErrors || isValidating}
						>
							Load
						</Button>
						<Button appearance="subtle" onClick={handleClose}>
							Cancel
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
