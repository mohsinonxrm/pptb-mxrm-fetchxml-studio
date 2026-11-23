/**
 * Undo and Redo for the tree toolbar's buttons and the keys (#192): runs the step, then announces it
 * politely ("Undone: Delete filter"), so screen reader users hear what changed. `announcement` goes
 * into a LiveRegion that stays mounted.
 *
 * The announcement waits until the presses stop (ANNOUNCE_DELAY_MS), so holding Ctrl+Z reads the
 * last step instead of a queue of them. It's cleared on each press, so the same text read twice in a
 * row ("Undone: Add filter" for two filters) is still a change the screen reader reads.
 *
 * The functions keep their identity and read the latest history, so a key listener can hold them.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import type { BuilderHistoryApi } from "../../features/fetchxml/state/builderContext";

export const ANNOUNCE_DELAY_MS = 300;

export interface UndoRedo {
	undo: () => void;
	redo: () => void;
	/** For a polite LiveRegion: "Undone: {step}" / "Redone: {step}", or "" */
	announcement: string;
}

export function useUndoRedo(history: BuilderHistoryApi): UndoRedo {
	const [announcement, setAnnouncement] = useState("");
	const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
	useEffect(() => () => clearTimeout(timer.current), []);
	// A layout effect: a second key press right after the first must see the step the first one left
	const latest = useRef(history);
	useLayoutEffect(() => {
		latest.current = history;
	});

	const announce = useCallback((message: string) => {
		clearTimeout(timer.current);
		setAnnouncement("");
		timer.current = setTimeout(() => setAnnouncement(message), ANNOUNCE_DELAY_MS);
	}, []);

	const undo = useCallback(() => {
		// With nothing to undo, undo() changes nothing and there's no label to announce
		const { undoLabel, undo: run } = latest.current;
		run();
		if (undoLabel) announce(`Undone: ${undoLabel}`);
	}, [announce]);

	const redo = useCallback(() => {
		const { redoLabel, redo: run } = latest.current;
		run();
		if (redoLabel) announce(`Redone: ${redoLabel}`);
	}, [announce]);

	return { undo, redo, announcement };
}
