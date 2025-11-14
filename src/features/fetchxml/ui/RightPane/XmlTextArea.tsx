/**
 * Simple XML text area viewer (fallback when Monaco fails due to CSP)
 */

import { Button, makeStyles, Tooltip, Textarea } from "@fluentui/react-components";
import { Copy24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		height: "100%",
		position: "relative",
		padding: "8px",
	},
	textarea: {
		flex: 1,
		fontFamily: "'Consolas', 'Monaco', 'Courier New', monospace",
		fontSize: "13px",
		lineHeight: "1.5",
	},
	copyButton: {
		position: "absolute",
		top: "16px",
		right: "16px",
		zIndex: 10,
	},
});

interface XmlTextAreaProps {
	xml: string;
}

export function XmlTextArea({ xml }: XmlTextAreaProps) {
	const styles = useStyles();

	const handleCopy = async () => {
		try {
			await navigator.clipboard.writeText(xml);
		} catch (err) {
			console.error("Failed to copy:", err);
		}
	};

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
			<Textarea
				className={styles.textarea}
				value={xml}
				readOnly
				resize="none"
				style={{ height: "100%" }}
			/>
		</div>
	);
}
