/**
 * Undo and Redo for the tree toolbar's buttons and the keys (#192): runs the step, then announces it
 * politely ("Undone: Delete filter") through the app's announcer (#186), so screen reader users hear
 * what changed.
 *
 * The announcement waits until the presses stop (ANNOUNCE_DELAY_MS), so holding Ctrl+Z reads the
 * last step instead of a queue of them. The announcer clears its text before each reading, so the
 * same text twice in a row ("Undone: Add filter" for two filters) is still read, and clears it again
 * a while later.
 *
 * The functions keep their identity and read the latest history, so a key listener can hold them.
 */
import { useCallback, useEffect, useLayoutEffect, useRef } from "react";
import type { BuilderHistoryApi } from "../../features/fetchxml/state/builderContext";
import { useAnnouncer } from "../../shared/feedback/useAnnouncer";

export const ANNOUNCE_DELAY_MS = 300;

export interface UndoRedo {
	undo: () => void;
	redo: () => void;
}

export function useUndoRedo(history: BuilderHistoryApi): UndoRedo {
	const announceNow = useAnnouncer();
	const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
	useEffect(() => () => clearTimeout(timer.current), []);
	// A layout effect: a second key press right after the first must see the step the first one left
	const latest = useRef({ history, announceNow });
	useLayoutEffect(() => {
		latest.current = { history, announceNow };
	});

	const announce = useCallback((message: string) => {
		clearTimeout(timer.current);
		timer.current = setTimeout(
			() => latest.current.announceNow(message, { batchId: "undo-redo" }),
			ANNOUNCE_DELAY_MS,
		);
	}, []);

	const undo = useCallback(() => {
		// With nothing to undo, undo() changes nothing and there's no label to announce
		const { undoLabel, undo: run } = latest.current.history;
		run();
		if (undoLabel) announce(`Undone: ${undoLabel}`);
	}, [announce]);

	const redo = useCallback(() => {
		const { redoLabel, redo: run } = latest.current.history;
		run();
		if (redoLabel) announce(`Redone: ${redoLabel}`);
	}, [announce]);

	return { undo, redo };
}
