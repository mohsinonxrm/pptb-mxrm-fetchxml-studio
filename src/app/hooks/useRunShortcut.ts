/**
 * Ctrl+Enter (Cmd+Enter on a Mac) runs the query from anywhere in the tool (#278): the tree, the
 * properties, the grid. It's the "run" entry of the shortcut table (#186), which extends this to the
 * other global shortcuts.
 *
 * It doesn't run:
 *  - in Monaco, where Ctrl+Enter inserts a line until #186 binds it there;
 *  - while a dialog or a menu is open: the shortcut belongs to what's in front;
 *  - when another handler already handled the key (defaultPrevented);
 *  - when Run can't be pressed (it has a reason).
 */
import { useEffect, useRef } from "react";
import { matches } from "../shortcuts/registry";

const BLOCKING_POPUPS = '[role="dialog"], [role="alertdialog"], [role="menu"]';

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
			if (document.querySelector(BLOCKING_POPUPS)) return;
			if (latest.current.disabledReason) return;
			event.preventDefault();
			latest.current.run();
		};
		window.addEventListener("keydown", onKeyDown);
		return () => window.removeEventListener("keydown", onKeyDown);
	}, []);
}
