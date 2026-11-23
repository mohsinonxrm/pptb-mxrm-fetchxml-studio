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
			for (const area of latest.current())
				if (area.element?.contains(target)) lastFocused.current.set(area.element, target);
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
		// From outside every area, Ctrl+F6 starts at the first and Ctrl+Shift+F6 at the last
		const index =
			current === -1
				? step === 1
					? 0
					: areas.length - 1
				: (current + step + areas.length) % areas.length;
		const area = areas[index];
		const remembered = lastFocused.current.get(area.element);
		if (remembered && area.element.contains(remembered) && isShown(remembered)) {
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
