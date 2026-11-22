/**
 * Ctrl+Z undoes and Ctrl+Y or Ctrl+Shift+Z redo the builder's last step (#192), from the tree, the
 * grid, a button or the page: the "undo" and "redo" entries of the shortcut table (#186), which
 * folds this listener into its one global handler later.
 *
 * It doesn't act (owner decision on #192):
 *  - in a text field, contenteditable or Monaco, which keep their own undo;
 *  - while a dialog or a menu is open: the keys belong to what's in front;
 *  - when something else already handled the key (`defaultPrevented`).
 *
 * The listener is on the window in the bubble phase, so a focused control's own key handling runs
 * first. Held down, the keys repeat, as undo does elsewhere.
 *
 * Undoing from the tree can remove the focused row (undoing an add): focus would fall to the page,
 * so `restoreFocus` puts it back in the tree, on the selected row.
 */
import { useEffect, useLayoutEffect, useRef } from "react";
import { isEditableTarget, matches } from "../shortcuts/registry";
import { hasOpenPopup } from "./useRunShortcut";

export interface UndoShortcutHandlers {
	undo: () => void;
	redo: () => void;
	/** Called after a key undo or redo from the tree when the focused row is gone */
	restoreFocus?: () => void;
}

const focusIsLost = () => !document.activeElement || document.activeElement === document.body;

export function useUndoShortcuts(handlers: UndoShortcutHandlers): void {
	const latest = useRef(handlers);
	useLayoutEffect(() => {
		latest.current = handlers;
	});

	useEffect(() => {
		const onKeyDown = (event: KeyboardEvent) => {
			if (event.defaultPrevented || event.isComposing) return;
			const isUndo = matches(event, "undo");
			if (!isUndo && !matches(event, "redo")) return;
			if (isEditableTarget(event.target)) return;
			if (hasOpenPopup()) return;
			event.preventDefault();
			const fromTree = !!(event.target as Element | null)?.closest?.('[role="tree"]');
			if (isUndo) latest.current.undo();
			else latest.current.redo();
			if (!fromTree) return;
			// After the render, and after the tree opens the restored node's branches
			requestAnimationFrame(() =>
				requestAnimationFrame(() => {
					if (focusIsLost()) latest.current.restoreFocus?.();
				}),
			);
		};
		window.addEventListener("keydown", onKeyDown);
		return () => window.removeEventListener("keydown", onKeyDown);
	}, []);
}
