/**
 * Focus goes back to whatever opened a dialog or drawer once it closes (#186, owner decision on
 * #186 in wave 1): the Properties field that opened "Turn off aggregate?", the tree row whose menu
 * opened Select attributes, the row where Ctrl+S was pressed.
 *
 * Fluent's own restore goes to the last control marked to receive it (toolbar buttons are), which
 * after a dialog opened from the tree or Properties is a toolbar button clicked earlier. So this
 * hook remembers the last element focused outside any popup, and when every modal dialog, menu and
 * trapped popover has closed, focuses it again; if that control is gone (a tree row's own Delete
 * button), the row or cell that held it. A hidden one can't take focus, so focus then stays put.
 *
 * Popups in a row (a menu, then the dialog its item opens; "Load into the tree?" before Save) count
 * as one: it waits two frames after the last one closes. Non-modal surfaces (the Properties drawer)
 * aren't popups (`hasOpenPopup`).
 */
import { useEffect } from "react";
import { hasOpenPopup } from "./useShortcuts";

const POPUP_ROOTS = '[role="dialog"], [role="alertdialog"], [role="menu"], [role="listbox"]';
/** A row or cell that holds the control: where focus goes when the control itself is gone */
const HOLDERS = '[role="treeitem"], [role="gridcell"], [role="row"]';

/**
 * Inside a popup that's in front: its focus isn't the place to come back to. The same rules as
 * hasOpenPopup: a non-modal surface (the Properties drawer) and Monaco's Find box don't count, so a
 * field in the drawer is a place to come back to.
 */
function inPopup(el: Element): boolean {
	const root = el.closest(POPUP_ROOTS);
	return !!root && root.getAttribute("aria-modal") !== "false" && !root.closest(".monaco-editor");
}

export function useFocusReturn(doc: Document = document): void {
	useEffect(() => {
		/** The last element focused outside every popup, and the one to return to */
		let lastOutside: HTMLElement | null = null;
		/** Its tree row or grid cell, found while it's still in the page */
		let lastOutsideHolder: HTMLElement | null = null;
		let returnTo: HTMLElement | null = null;
		/** The tree row or grid cell holding returnTo: a row's own buttons unmount once they blur */
		let returnToHolder: HTMLElement | null = null;
		let anyOpen = hasOpenPopup(doc);
		let settling = false;

		const onFocusIn = (event: FocusEvent) => {
			const target = event.target as HTMLElement | null;
			if (target && target !== doc.body && !inPopup(target)) {
				lastOutside = target;
				lastOutsideHolder =
					(target.parentElement?.closest(HOLDERS) as HTMLElement | null) ?? null;
			}
		};
		// Focus that falls to the page (a click on empty space) leaves nothing to come back to
		const onFocusOut = () =>
			queueMicrotask(() => {
				const active = doc.activeElement;
				if ((!active || active === doc.body) && !hasOpenPopup(doc)) {
					lastOutside = null;
					lastOutsideHolder = null;
				}
			});

		const check = () => {
			const nowOpen = hasOpenPopup(doc);
			// A popup appears where none was: what had focus is what opened it. A popup that follows
			// another within the settle (a menu, then its dialog) keeps the first one's target.
			if (nowOpen && !anyOpen && !settling) {
				returnTo = lastOutside;
				returnToHolder = lastOutsideHolder;
			}
			anyOpen = nowOpen;
			if (nowOpen || settling || !returnTo) return;
			settling = true;
			requestAnimationFrame(() =>
				requestAnimationFrame(() => {
					settling = false;
					if (hasOpenPopup(doc)) return;
					const target = returnTo?.isConnected ? returnTo : returnToHolder;
					returnTo = null;
					returnToHolder = null;
					if (target && doc.activeElement !== target) target.focus();
				}),
			);
		};
		// The body changes constantly (grid scrolling, Monaco's cursor): look only at changes that add,
		// remove or touch a popup, or at any change while one is open. Right away, so a short-lived
		// dialog isn't missed.
		const touchesPopup = (record: MutationRecord) =>
			(record.type === "childList"
				? [...record.addedNodes, ...record.removedNodes]
				: [record.target]
			).some(
				(node) =>
					node instanceof Element &&
					(node.matches(POPUP_ROOTS) || !!node.querySelector(POPUP_ROOTS)),
			);

		doc.addEventListener("focusin", onFocusIn, true);
		doc.addEventListener("focusout", onFocusOut, true);
		const observer = new MutationObserver((records) => {
			if (anyOpen || records.some(touchesPopup)) check();
		});
		observer.observe(doc.body, {
			childList: true,
			subtree: true,
			attributes: true,
			attributeFilter: ["hidden", "style", "aria-modal"],
		});
		return () => {
			doc.removeEventListener("focusin", onFocusIn, true);
			doc.removeEventListener("focusout", onFocusOut, true);
			observer.disconnect();
		};
	}, [doc]);
}
