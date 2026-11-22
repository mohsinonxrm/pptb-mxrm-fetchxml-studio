/**
 * Undo and Redo for the tree toolbar's buttons and the keys (#192): runs the step, then announces it
 * politely ("Undone: Delete filter"), so screen reader users hear what changed. The announcement goes
 * through Fluent's AriaLiveAnnouncer, mounted once at the top of the app.
 *
 * The functions keep their identity and read the latest history, so a key listener can hold them.
 */
import { useCallback, useLayoutEffect, useRef } from "react";
import { useAnnounce } from "@fluentui/react-components";
import type { BuilderHistoryApi } from "../../features/fetchxml/state/builderContext";

/** One batch, so pressing Ctrl+Z quickly announces the last step instead of a queue of them */
const BATCH = "fxs-undo-redo";

export function useUndoRedo(history: BuilderHistoryApi): { undo: () => void; redo: () => void } {
	const { announce } = useAnnounce();
	// A layout effect: a second key press right after the first must see the step the first one left
	const latest = useRef(history);
	useLayoutEffect(() => {
		latest.current = history;
	});

	const undo = useCallback(() => {
		// With nothing to undo, undo() changes nothing and there's no label to announce
		const { undoLabel, undo: run } = latest.current;
		run();
		if (undoLabel) announce(`Undone: ${undoLabel}`, { polite: true, batchId: BATCH });
	}, [announce]);

	const redo = useCallback(() => {
		const { redoLabel, redo: run } = latest.current;
		run();
		if (redoLabel) announce(`Redone: ${redoLabel}`, { polite: true, batchId: BATCH });
	}, [announce]);

	return { undo, redo };
}
