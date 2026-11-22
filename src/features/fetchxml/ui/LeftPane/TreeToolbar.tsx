/**
 * The tree toolbar (#192, design D): Undo and Redo above the query tree, in the docked pane and the
 * tree peek. A shared overflow toolbar (#232), so #283 can add Add, Delete and More to it later.
 *
 * Icon-only buttons whose Tooltip names the step ("Undo: Delete filter (Ctrl+Z)"). With nothing to
 * undo or redo they stay focusable, with the reason ("Undo: Nothing to undo").
 *
 * Hidden before there's a query (the first-run picker), as in the design, but not while there's a
 * step to redo: undoing the first entity pick leaves no query, and Redo has to stay reachable.
 */
import {
	ArrowRedo20Filled,
	ArrowRedo20Regular,
	ArrowUndo20Filled,
	ArrowUndo20Regular,
	bundleIcon,
} from "@fluentui/react-icons";
import { OverflowToolbar, type BarCommand } from "../../../../shared/components/commandBar";
import { formatKeys, getEntry } from "../../../../app/shortcuts/registry";

const UndoIcon = bundleIcon(ArrowUndo20Filled, ArrowUndo20Regular);
const RedoIcon = bundleIcon(ArrowRedo20Filled, ArrowRedo20Regular);

export interface TreeToolbarProps {
	/** Is there a query tree? */
	hasQuery: boolean;
	canUndo: boolean;
	canRedo: boolean;
	undoLabel: string | null;
	redoLabel: string | null;
	onUndo: () => void;
	onRedo: () => void;
}

/** "Undo: Delete filter (Ctrl+Z)", or plain "Undo" when there's nothing to undo (the reason follows) */
function commandLabel(verb: string, step: string | null, keys: string): string {
	return step ? `${verb}: ${step} (${keys})` : verb;
}

export function TreeToolbar({
	hasQuery,
	canUndo,
	canRedo,
	undoLabel,
	redoLabel,
	onUndo,
	onRedo,
}: TreeToolbarProps) {
	if (!hasQuery && !canUndo && !canRedo) return null;
	const undoEntry = getEntry("undo");
	const redoEntry = getEntry("redo");
	const commands: BarCommand[] = [
		{
			id: "undo",
			kind: "button",
			group: "history",
			label: commandLabel("Undo", canUndo ? undoLabel : null, formatKeys(undoEntry)[0]),
			icon: UndoIcon,
			iconOnly: true,
			disabledReason: canUndo ? undefined : "Nothing to undo",
			keyShortcuts: undoEntry.aria,
			onClick: onUndo,
		},
		{
			id: "redo",
			kind: "button",
			group: "history",
			label: commandLabel("Redo", canRedo ? redoLabel : null, formatKeys(redoEntry)[0]),
			icon: RedoIcon,
			iconOnly: true,
			disabledReason: canRedo ? undefined : "Nothing to redo",
			keyShortcuts: redoEntry.aria,
			onClick: onRedo,
		},
	];
	return (
		<OverflowToolbar
			ariaLabel="Query tree commands"
			moreLabel="More tree commands"
			size="small"
			commands={commands}
		/>
	);
}
