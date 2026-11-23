/**
 * Ctrl+F6 / Ctrl+Shift+F6 between the tool's areas (#186, owner decision Q5, as in model-driven
 * apps): app bar → builder → output and round again. Each area remembers the last element that had
 * focus in it, and going back there returns to it; the first time, or when that element is gone or
 * hidden, focus goes to the area's first focusable element (`focusFirst`).
 *
 * Areas that aren't shown (the builder pane while it's collapsed, say) are skipped: `getAreas`
 * returns them as null, or they (or an ancestor) are display: none or hidden.
 */
import { useCallback, useEffect, useRef } from "react";

export interface FocusArea {
	element: HTMLElement | null;
	/** Elements that belong to the area too, without being where focus goes (the rail beside a peek) */
	alsoContains?: (HTMLElement | null)[];
	/** Focus the area's natural first stop (the selected tree item, Run…) */
	focusFirst: (element: HTMLElement) => void;
}

/** Connected, and neither it nor an ancestor is display: none or hidden */
function isShown(el: HTMLElement): boolean {
	if (!el.isConnected) return false;
	for (let node: HTMLElement | null = el; node; node = node.parentElement) {
		if (node.hidden || getComputedStyle(node).display === "none") return false;
	}
	return true;
}
const follows = (el: Node, ref: Node) =>
	!!(ref.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING) && !el.contains(ref);
const precedes = (el: Node, ref: Node) =>
	!!(ref.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_PRECEDING) && !el.contains(ref);
function lastIndex<T>(items: T[], test: (item: T) => boolean, fallback: number): number {
	for (let i = items.length - 1; i >= 0; i--) if (test(items[i])) return i;
	return fallback;
}
const holds = (area: FocusArea, target: Node) =>
	[area.element, ...(area.alsoContains ?? [])].some((el) => !!el?.contains(target));

export function useFocusAreas(getAreas: () => FocusArea[]): {
	next: () => void;
	previous: () => void;
} {
	const latest = useRef(getAreas);
	useEffect(() => {
		latest.current = getAreas;
	});
	/** The last focused element in each area, by area element */
	const lastFocused = useRef(new WeakMap<HTMLElement, HTMLElement>());

	useEffect(() => {
		const onFocusIn = (event: FocusEvent) => {
			const target = event.target as HTMLElement | null;
			if (!target) return;
			// The first area that holds it: the Properties drawer is the builder's, though it sits in
			// the output
			const area = latest.current().find((a) => a.element && holds(a, target));
			if (area?.element) lastFocused.current.set(area.element, target);
		};
		document.addEventListener("focusin", onFocusIn);
		return () => document.removeEventListener("focusin", onFocusIn);
	}, []);

	const move = useCallback((step: 1 | -1) => {
		const areas = latest
			.current()
			.filter(
				(a): a is FocusArea & { element: HTMLElement } => !!a.element && isShown(a.element),
			);
		if (areas.length === 0) return;
		const active = document.activeElement as HTMLElement | null;
		const current = areas.findIndex((a) => !!active && holds(a, active));
		// From outside every area (a splitter between two, a toast), the next area after it in the
		// page, or the previous one before it; past the end, round to the first or the last
		const index =
			current !== -1
				? (current + step + areas.length) % areas.length
				: step === 1
					? Math.max(
							0,
							areas.findIndex((a) => !!active && follows(a.element, active)),
						)
					: lastIndex(
							areas,
							(a) => !!active && precedes(a.element, active),
							areas.length - 1,
						);
		const area = areas[index];
		const remembered = lastFocused.current.get(area.element);
		if (remembered && holds(area, remembered) && isShown(remembered)) {
			remembered.focus();
			if (document.activeElement === remembered) return;
		}
		area.focusFirst(area.element);
	}, []);

	return {
		next: useCallback(() => move(1), [move]),
		previous: useCallback(() => move(-1), [move]),
	};
}
