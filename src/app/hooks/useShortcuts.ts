/**
 * The tool's one keyboard shortcut handler (#186): the "global" entries of the shortcut table
 * (`src/app/shortcuts/registry.ts`). It replaces #278's Ctrl+Enter and #192's Ctrl+Z / Ctrl+Y
 * listeners. F2 isn't here: the Toaster focuses the newest toast itself (NotifyProvider).
 *
 * Two listeners on the window:
 *  - Capture phase, for keys that type no text and so also work in fields and the FetchXML editor:
 *    Ctrl+Enter (Run), Ctrl+S (Save), Ctrl+F6 / Ctrl+Shift+F6 (areas). Handled keys stop there, so
 *    Enter doesn't also expand a tree branch. In the FetchXML editor, Ctrl+Enter is Monaco's own
 *    "Run query" action (RUNS_ITSELF_ATTRIBUTE), and in a Find box it stays Find's.
 *  - Bubble phase, for keys a focused control may have its own use for: Ctrl+Z / Ctrl+Y /
 *    Ctrl+Shift+Z (Undo, Redo) and Ctrl+/ (the shortcuts list). They act only outside text fields,
 *    contenteditable and Monaco (which keep their own undo and comment toggle), and only when
 *    nothing handled the key first.
 *
 * Nothing acts while a modal dialog or a menu is open: the keys belong to what's in front. The
 * non-modal Properties drawer doesn't count (#385). Ctrl+Enter also leaves a focused button or menu
 * item alone (Enter is that control's own action). Held down, or while Run has a reason (running,
 * say), Ctrl+Enter runs nothing but is still taken, so it doesn't toggle a tree branch either.
 *
 * Ctrl+S and Ctrl+/ open a dialog from wherever focus is; useFocusReturn brings focus back there
 * once it closes, as for every dialog.
 *
 * An undo can remove the focused element (undoing an add removes its tree row, its properties or its
 * grid column). Focus would fall to the page, so `restoreFocus` puts it somewhere useful.
 */
import { useEffect, useLayoutEffect, useRef } from "react";
import { RUNS_ITSELF_ATTRIBUTE, isEditableTarget, matches } from "../shortcuts/registry";

const POPUPS = '[role="dialog"], [role="alertdialog"], [role="menu"]';
const OWN_ENTER = 'button, a[href], [role="button"], [role^="menuitem"]';
/** Ctrl+Enter belongs to these: the FetchXML editor's own Run action, and Monaco's Find box */
const OWN_CTRL_ENTER = `[${RUNS_ITSELF_ATTRIBUTE}], .find-widget`;

/**
 * An open dialog or menu in front. Monaco's Find widget (role="dialog") doesn't count: it stays in
 * the editor after it closes, and the editor keeps its own keys anyway. Hidden ones don't either.
 * (Not aria-hidden: Fluent's modal surfaces can sit in an aria-hidden container while they open.)
 * A non-modal dialog (aria-modal="false", like the Properties drawer) isn't in front of anything, so
 * it doesn't count (#385).
 */
export function hasOpenPopup(doc: Document = document): boolean {
	return [...doc.querySelectorAll<HTMLElement>(POPUPS)].some(
		(el) =>
			el.getAttribute("aria-modal") !== "false" &&
			!el.closest("[hidden], .monaco-editor") &&
			doc.defaultView?.getComputedStyle(el).display !== "none",
	);
}

/** A dropdown or menu button with its list open: the keys are its until it closes */
function hasOpenList(target: EventTarget | null): boolean {
	const el = target as Element | null;
	if (!el || typeof el.getAttribute !== "function") return false;
	return (
		el.getAttribute("aria-expanded") === "true" &&
		(el.getAttribute("role") === "combobox" || el.hasAttribute("aria-haspopup"))
	);
}

const focusIsLost = () => !document.activeElement || document.activeElement === document.body;

export interface ShortcutHandlers {
	/** Run (Ctrl+Enter); nothing happens while `runDisabledReason` is set */
	run: () => void;
	runDisabledReason?: string;
	/** A key that can't act now, with the reason (Run has one): announced (#186 PR 3) */
	onBlocked?: (reason: string) => void;
	/** Save (Ctrl+S): the same as selecting Save */
	save: () => void;
	undo: () => void;
	redo: () => void;
	/** Called after a key undo or redo took focus from an element and left it on the page */
	restoreFocus?: () => void;
	/** Ctrl+/: the Keyboard shortcuts dialog */
	showShortcuts: () => void;
	/** Ctrl+F6 / Ctrl+Shift+F6: the next or previous area (app bar, builder, output) */
	nextArea: () => void;
	previousArea: () => void;
}

export function useShortcuts(handlers: ShortcutHandlers): void {
	// The listeners stay put; they call the latest handlers. A layout effect: a second key press
	// right after the first must see what the first one left.
	const latest = useRef(handlers);
	useLayoutEffect(() => {
		latest.current = handlers;
	});

	useEffect(() => {
		/** Keys that type no text: in fields and the editor too, stopped once handled */
		const onCapture = (event: KeyboardEvent) => {
			if (event.defaultPrevented || event.isComposing) return;
			const h = latest.current;
			const act = (fn: () => void) => {
				event.preventDefault();
				event.stopPropagation();
				fn();
			};
			if (matches(event, "run")) {
				const target = event.target as Element | null;
				if (target?.closest?.(OWN_ENTER) || target?.closest?.(OWN_CTRL_ENTER)) return;
				if (hasOpenPopup()) return;
				act(
					event.repeat
						? () => {}
						: h.runDisabledReason
							? () => h.onBlocked?.(h.runDisabledReason!)
							: h.run,
				);
			} else if (matches(event, "save")) {
				if (event.repeat || hasOpenPopup()) return;
				act(h.save);
			} else if (matches(event, "nextArea")) {
				if (!hasOpenPopup()) act(h.nextArea);
			} else if (matches(event, "previousArea")) {
				if (!hasOpenPopup()) act(h.previousArea);
			}
		};

		/** Keys a control may use itself: outside text fields, after the control had its turn */
		const onBubble = (event: KeyboardEvent) => {
			if (event.defaultPrevented || event.isComposing) return;
			const isUndo = matches(event, "undo");
			const isRedo = !isUndo && matches(event, "redo");
			const isList = !isUndo && !isRedo && matches(event, "shortcuts");
			if (!isUndo && !isRedo && !isList) return;
			if (isEditableTarget(event.target) || hasOpenList(event.target)) return;
			if (hasOpenPopup()) return;
			event.preventDefault();
			const h = latest.current;
			if (isList) {
				if (!event.repeat) h.showShortcuts();
				return;
			}
			const hadFocus = !focusIsLost();
			if (isUndo) h.undo();
			else h.redo();
			// After the render, and after the tree opens the restored node's branches
			requestAnimationFrame(() =>
				requestAnimationFrame(() => {
					if (hadFocus && focusIsLost()) latest.current.restoreFocus?.();
				}),
			);
		};

		window.addEventListener("keydown", onCapture, true);
		window.addEventListener("keydown", onBubble);
		return () => {
			window.removeEventListener("keydown", onCapture, true);
			window.removeEventListener("keydown", onBubble);
		};
	}, []);
}
