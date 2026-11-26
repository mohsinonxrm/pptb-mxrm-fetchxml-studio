/**
 * Focus moves to an element when a condition becomes true (#279). A record dialog's primary button
 * goes away once the job starts: focus moves to the running Cancel, then to Close when the job ends,
 * so it never drops to the page.
 */
import { useEffect, useRef } from "react";

export function useFocusWhen<T extends HTMLElement = HTMLButtonElement>(when: boolean) {
	const ref = useRef<T>(null);
	useEffect(() => {
		if (when) ref.current?.focus();
	}, [when]);
	return ref;
}
