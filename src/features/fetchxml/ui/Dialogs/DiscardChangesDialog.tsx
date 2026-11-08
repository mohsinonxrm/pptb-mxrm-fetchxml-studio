/**
 * "Discard unsaved changes?" before New and Open view (#278, #192). Not wired yet: the app bar's
 * New and Open view will render it when the query has unsaved changes. Save and ... runs the
 * caller's save first (the caller continues afterwards); Discard changes continues without saving;
 * Cancel and Escape keep the query. While the save runs, Discard waits and Cancel stays available.
 */

import {
	Button,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	Spinner,
	Tooltip,
} from "@fluentui/react-components";

export interface DiscardChangesDialogProps {
	open: boolean;
	/** What the user is about to do */
	intent: "new" | "open";
	/** The loaded view, or null when the query isn't from a view */
	viewName: string | null;
	/** The view being opened (intent "open") */
	targetViewName?: string;
	/** Saves; the caller continues (new or open) once it succeeds */
	onSave: () => void | Promise<void>;
	onDiscard: () => void;
	onCancel: () => void;
	/** A save is running */
	saving?: boolean;
}

function discardChangesBody(
	intent: "new" | "open",
	viewName: string | null,
	targetViewName?: string,
): string {
	const lead =
		viewName === null
			? "Your query isn't saved."
			: `Your changes to "${viewName}" aren't saved.`;
	const effect =
		intent === "open"
			? `Opening "${targetViewName ?? ""}" replaces the tree, the FetchXML and the results.`
			: "A new query clears the tree, the FetchXML and the results.";
	return `${lead} ${effect}`;
}

export function DiscardChangesDialog({
	open,
	intent,
	viewName,
	targetViewName,
	onSave,
	onDiscard,
	onCancel,
	saving = false,
}: DiscardChangesDialogProps) {
	const saveLabel = intent === "open" ? "Save and open" : "Save and start new";
	const discard = (
		<Button
			appearance="secondary"
			disabledFocusable={saving}
			onClick={saving ? undefined : onDiscard}
		>
			Discard changes
		</Button>
	);
	return (
		<Dialog
			open={open}
			modalType="alert"
			onOpenChange={(_event, data) => {
				if (!data.open) onCancel();
			}}
		>
			<DialogSurface>
				<DialogBody>
					<DialogTitle>Discard unsaved changes?</DialogTitle>
					<DialogContent>
						{discardChangesBody(intent, viewName, targetViewName)}
					</DialogContent>
					<DialogActions>
						<Button
							appearance="primary"
							autoFocus
							disabledFocusable={saving}
							icon={saving ? <Spinner size="tiny" /> : undefined}
							onClick={saving ? undefined : () => void onSave()}
						>
							{saving ? "Saving…" : saveLabel}
						</Button>
						{saving ? (
							<Tooltip
								content="Wait for the save to finish"
								relationship="description"
							>
								{discard}
							</Tooltip>
						) : (
							discard
						)}
						<Button appearance="secondary" onClick={onCancel}>
							Cancel
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
