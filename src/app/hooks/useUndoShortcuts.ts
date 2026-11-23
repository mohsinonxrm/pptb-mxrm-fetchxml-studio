/**
 * Ctrl+Z undoes and Ctrl+Y or Ctrl+Shift+Z redo the builder's last step (#192), from the tree, the
 * grid, a button or the page: the "undo" and "redo" entries of the shortcut table (#186), which
 * folds this listener into its one global handler later.
 *
 * It doesn't act (owner decision on #192):
 *  - in a text field, contenteditable or Monaco, which keep their own undo;
 *  - while a modal dialog or a menu is open: the keys belong to what's in front. The non-modal
 *    Properties drawer isn't in front of anything, so the keys work beside it and in it (#385);
 *  - on a dropdown or menu button whose list is open;
 *  - when something else already handled the key (`defaultPrevented`).
 *
 * The listener is on the window in the bubble phase, so a focused control's own key handling runs
 * first. Held down, the keys repeat, as undo does elsewhere.
 *
 * An undo can remove the focused element (undoing an add removes its tree row, its properties or its
 * grid column). Focus would fall to the page, so `restoreFocus` puts it somewhere useful.
 */
import { useEffect, useLayoutEffect, useRef } from "react";
import { isEditableTarget, matches } from "../shortcuts/registry";
import { hasOpenPopup } from "./useRunShortcut";

export interface UndoShortcutHandlers {
	undo: () => void;
	redo: () => void;
	/** Called after a key undo or redo took focus from an element and left it on the page */
	restoreFocus?: () => void;
}

const focusIsLost = () => !document.activeElement || document.activeElement === document.body;

/** A dropdown or menu button with its list open: the keys are its until it closes */
function hasOpenList(target: EventTarget | null): boolean {
	const el = target as Element | null;
	if (!el || typeof el.getAttribute !== "function") return false;
	return (
		el.getAttribute("aria-expanded") === "true" &&
		(el.getAttribute("role") === "combobox" || el.hasAttribute("aria-haspopup"))
	);
}

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
			if (isEditableTarget(event.target) || hasOpenList(event.target)) return;
			if (hasOpenPopup()) return;
			event.preventDefault();
			const hadFocus = !focusIsLost();
			if (isUndo) latest.current.undo();
			else latest.current.redo();
			// After the render, and after the tree opens the restored node's branches
			requestAnimationFrame(() =>
				requestAnimationFrame(() => {
					if (hadFocus && focusIsLost()) latest.current.restoreFocus?.();
				}),
			);
		};
		window.addEventListener("keydown", onKeyDown);
		return () => window.removeEventListener("keydown", onKeyDown);
	}, []);
}
