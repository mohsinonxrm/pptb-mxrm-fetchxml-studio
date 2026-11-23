/**
 * Ctrl+Enter (Cmd+Enter on a Mac) runs the query from anywhere in the tool (#278): the tree, the
 * properties, the pickers, the grid. It's the "run" entry of the shortcut table (#186), which
 * extends this to the other global shortcuts.
 *
 * The listener runs in the capture phase and, when it runs the query, stops the key there: Enter
 * would otherwise also expand or collapse a tree branch, or open a picker's list.
 *
 * It doesn't run:
 *  - in Monaco, where Ctrl+Enter inserts a line until #186 binds it there;
 *  - on a button or a menu item: Enter is that control's own action ("Run anyway", say);
 *  - while a modal dialog or a menu is open: the shortcut belongs to what's in front (the
 *    non-modal Properties drawer doesn't count, #385);
 *  - when Run can't be pressed (it has a reason).
 */
import { useEffect, useRef } from "react";
import { matches } from "../shortcuts/registry";

const POPUPS = '[role="dialog"], [role="alertdialog"], [role="menu"]';
const OWN_ENTER = 'button, a[href], [role="button"], [role^="menuitem"]';

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

export function useRunShortcut(run: () => void, disabledReason: string | undefined): void {
	// The listener stays put; it calls the latest run and reads the latest reason
	const latest = useRef({ run, disabledReason });
	useEffect(() => {
		latest.current = { run, disabledReason };
	});

	useEffect(() => {
		const onKeyDown = (event: KeyboardEvent) => {
			if (event.defaultPrevented || event.repeat || !matches(event, "run")) return;
			const target = event.target as Element | null;
			if (target?.closest?.(".monaco-editor")) return;
			if (target?.closest?.(OWN_ENTER)) return;
			if (hasOpenPopup()) return;
			if (latest.current.disabledReason) return;
			event.preventDefault();
			event.stopPropagation();
			latest.current.run();
		};
		window.addEventListener("keydown", onKeyDown, true);
		return () => window.removeEventListener("keydown", onKeyDown, true);
	}, []);
}
