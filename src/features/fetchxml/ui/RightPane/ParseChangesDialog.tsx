/**
 * Confirmation shown before FetchXML is loaded into the tree when the tree would change the query
 * (#89): Parse to Tree, and the editor-mode actions that parse first (Execute, Save, Send, #102).
 */
import {
	Body1,
	Body1Strong,
	Button,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import type { ParseWarning } from "../../model/fetchxmlParser";
import { ParseIssueList } from "./ParseIssueList";
import type { EditorSyncAction } from "./editorSync";

const useStyles = makeStyles({
	content: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalM,
	},
});

const COPY: Record<"parse" | EditorSyncAction, { title: string; intro: string; confirm: string }> = {
	parse: {
		title: "Parse to Tree will change this query",
		intro: "The tree can't represent everything in this FetchXML. Parsing will make these changes:",
		confirm: "Parse anyway",
	},
	execute: {
		title: "Executing will change this query",
		intro:
			"In editor mode, Execute loads your FetchXML into the tree first. The tree can't represent everything in it, so these changes will be made:",
		confirm: "Load and execute",
	},
	save: {
		title: "Saving will change this query",
		intro:
			"In editor mode, saving loads your FetchXML into the tree first. The tree can't represent everything in it, so these changes will be made:",
		confirm: "Load and save",
	},
	send: {
		title: "Sending will change this query",
		intro:
			"In editor mode, sending loads your FetchXML into the tree first. The tree can't represent everything in it, so these changes will be made:",
		confirm: "Load and send",
	},
};

interface ParseChangesDialogProps {
	/** Which action asked for the confirmation; null hides the dialog */
	action: "parse" | EditorSyncAction | null;
	changes: ParseWarning[];
	onConfirm: () => void;
	onCancel: () => void;
}

export function ParseChangesDialog({ action, changes, onConfirm, onCancel }: ParseChangesDialogProps) {
	const styles = useStyles();
	const copy = COPY[action ?? "parse"];
	const removed = changes.filter((c) => c.kind === "dropped");
	const corrected = changes.filter((c) => c.kind === "corrected");
	return (
		<Dialog
			modalType="alert"
			open={action !== null}
			onOpenChange={(_, data) => {
				if (!data.open) onCancel();
			}}
		>
			<DialogSurface>
				<DialogBody>
					<DialogTitle>{copy.title}</DialogTitle>
					<DialogContent className={styles.content}>
						<Body1>{copy.intro}</Body1>
						{removed.length > 0 && (
							<div>
								<Body1Strong>Removed</Body1Strong>
								<ParseIssueList items={removed} />
							</div>
						)}
						{corrected.length > 0 && (
							<div>
								<Body1Strong>Corrected</Body1Strong>
								<ParseIssueList items={corrected} />
							</div>
						)}
						<Body1>If you cancel, the editor keeps your FetchXML unchanged.</Body1>
					</DialogContent>
					<DialogActions>
						<Button appearance="secondary" onClick={onCancel}>
							Cancel
						</Button>
						<Button appearance="primary" onClick={onConfirm}>
							{copy.confirm}
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
