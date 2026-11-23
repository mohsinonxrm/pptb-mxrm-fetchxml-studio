/**
 * Focus goes back to whatever opened a dialog or drawer once it closes (#186, owner decision on
 * #186 in wave 1): the Properties field that opened "Turn off aggregate?", the tree row whose menu
 * opened Select attributes, the row where Ctrl+S was pressed.
 *
 * Fluent's own restore goes to the last control marked to receive it (toolbar buttons are), which
 * after a dialog opened from the tree or Properties is a toolbar button clicked earlier. So this
 * hook remembers the last element focused outside any popup, and when every modal dialog, menu and
 * trapped popover has closed, focuses it again if it's still in the page and shown.
 *
 * Popups in a row (a menu, then the dialog its item opens; "Load into the tree?" before Save) count
 * as one: it waits two frames after the last one closes. Non-modal surfaces (the Properties drawer)
 * aren't popups (`hasOpenPopup`).
 */
import { useEffect } from "react";
import { hasOpenPopup } from "./useShortcuts";

const POPUP_ROOTS = '[role="dialog"], [role="alertdialog"], [role="menu"], [role="listbox"]';

/** Inside a popup: its focus isn't the place to come back to */
const inPopup = (el: Element) => !!el.closest(POPUP_ROOTS);

function isShown(el: HTMLElement): boolean {
	if (!el.isConnected) return false;
	for (let node: HTMLElement | null = el; node; node = node.parentElement) {
		if (node.hidden || getComputedStyle(node).display === "none") return false;
	}
	return true;
}

export function useFocusReturn(doc: Document = document): void {
	useEffect(() => {
		/** The last element focused outside every popup, and the one to return to */
		let lastOutside: HTMLElement | null = null;
		let returnTo: HTMLElement | null = null;
		let anyOpen = hasOpenPopup(doc);
		let settling = false;

		const onFocusIn = (event: FocusEvent) => {
			const target = event.target as HTMLElement | null;
			if (target && target !== doc.body && !inPopup(target)) lastOutside = target;
		};

		const check = () => {
			const nowOpen = hasOpenPopup(doc);
			// A popup appears where none was: what had focus is what opened it. In a run (a menu,
			// then its dialog) focus is back on the menu's trigger in between, so that's kept.
			if (nowOpen && !anyOpen) returnTo = lastOutside;
			anyOpen = nowOpen;
			if (nowOpen || settling || !returnTo) return;
			settling = true;
			requestAnimationFrame(() =>
				requestAnimationFrame(() => {
					settling = false;
					if (hasOpenPopup(doc)) return;
					const target = returnTo;
					returnTo = null;
					if (target && isShown(target) && doc.activeElement !== target) target.focus();
				}),
			);
		};

		doc.addEventListener("focusin", onFocusIn, true);
		const observer = new MutationObserver(check);
		observer.observe(doc.body, {
			childList: true,
			subtree: true,
			attributes: true,
			attributeFilter: ["hidden", "style", "aria-modal"],
		});
		return () => {
			doc.removeEventListener("focusin", onFocusIn, true);
			observer.disconnect();
		};
	}, [doc]);
}
