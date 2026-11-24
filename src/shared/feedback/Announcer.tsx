/**
 * The tool's screen reader announcements (#186, owner decision 2026-11-22): one announcer with a
 * polite and an assertive live region, visually hidden and always mounted (a region added together
 * with its text isn't read).
 *
 * - `useAnnouncer()` announces: Undo / Redo ("Undone: …"), the run status ("50 rows loaded…"), a
 *   key's reason when it can't act.
 * - Fluent components announce through it too: the provider gives Fluent's `AnnounceProvider`, so a
 *   MessageBar announces itself when it appears, polite for info and assertive for warnings and
 *   errors, as Fluent intends. `SilentAnnouncements` turns that off for a part of the page
 *   (Properties, whose bars would re-announce on every node selection), and `AnnouncementScope`
 *   makes a tab's bars announce once per change rather than on every visit.
 *
 * Fluent's own AriaLiveAnnouncer isn't used: without `ariaNotify` its region is always assertive.
 *
 * Messages that arrive close together are read as one (FLUSH_MS), in order; a message with the same
 * `batchId` as a waiting one replaces it, so a burst ("Undone" pressed five times, Retrieve all
 * progress) reads its last step. The text is cleared before each new reading, so the same message
 * twice is read twice, and again after CLEAR_AFTER_MS, so a later read of the page doesn't meet it.
 */
import {
	useCallback,
	useEffect,
	useLayoutEffect,
	useMemo,
	useRef,
	useState,
	type ReactNode,
} from "react";
import { createPortal } from "react-dom";
import { AnnounceProvider, makeStyles, useAnnounce } from "@fluentui/react-components";
import { AnnouncerContext, CLEAR_AFTER_MS, FLUSH_MS, type Announce } from "./useAnnouncer";

const useStyles = makeStyles({
	hidden: {
		position: "absolute",
		width: "1px",
		height: "1px",
		margin: "-1px",
		padding: 0,
		overflow: "hidden",
		clip: "rect(0px, 0px, 0px, 0px)",
		whiteSpace: "nowrap",
		border: 0,
	},
});

/** One region's waiting messages, by batch, and its timers */
function useRegion() {
	const [text, setText] = useState("");
	const pending = useRef(new Map<string, string>());
	const timers = useRef<{
		flush?: ReturnType<typeof setTimeout>;
		clear?: ReturnType<typeof setTimeout>;
	}>({});
	const counter = useRef(0);
	useEffect(
		() => () => {
			clearTimeout(timers.current.flush);
			clearTimeout(timers.current.clear);
		},
		[],
	);
	/** A waiting message of this batch isn't read after all (a run failed before "Running…" was) */
	const cancel = useCallback((batchId: string) => {
		pending.current.delete(batchId);
	}, []);
	const add = useCallback((message: string, batchId?: string) => {
		const key = batchId ?? `#${counter.current++}`;
		// A replaced batch moves to the end: it's the newest news
		pending.current.delete(key);
		pending.current.set(key, message);
		if (timers.current.flush) return;
		clearTimeout(timers.current.clear);
		setText("");
		timers.current.flush = setTimeout(() => {
			timers.current.flush = undefined;
			const joined = [...pending.current.values()].join(" ");
			pending.current.clear();
			if (!joined) return;
			setText(joined);
			timers.current.clear = setTimeout(() => setText(""), CLEAR_AFTER_MS);
		}, FLUSH_MS);
	}, []);
	return { text, add, cancel };
}

export function AnnouncerProvider({ children }: { children: ReactNode }) {
	const styles = useStyles();
	const polite = useRegion();
	const assertive = useRegion();
	const { add: addPolite, cancel: cancelPolite } = polite;
	const { add: addAssertive, cancel: cancelAssertive } = assertive;
	const announce = useCallback<Announce>(
		(message, options) => {
			const text = message.trim();
			// An empty message cancels its batch's waiting one
			if (!text) {
				if (options?.batchId) {
					cancelPolite(options.batchId);
					cancelAssertive(options.batchId);
				}
				return;
			}
			(options?.politeness === "assertive" ? addAssertive : addPolite)(
				text,
				options?.batchId,
			);
		},
		[addPolite, addAssertive, cancelPolite, cancelAssertive],
	);
	// Fluent's MessageBar asks for "alert" when it's a warning or an error
	const fluent = useMemo(
		() => ({
			announce: (message: string, options?: { alert?: boolean; batchId?: string }) =>
				announce(message, {
					politeness: options?.alert ? "assertive" : "polite",
					batchId: options?.batchId,
				}),
		}),
		[announce],
	);
	// The regions are top-level children of the body, marked never-hide: a modal dialog's focus trap
	// sets aria-hidden on everything else at the top level (Tabster), which would silence them while a
	// dialog is open. Fluent's own announcer does the same.
	const regions = (
		<>
			<div
				role="status"
				aria-live="polite"
				aria-atomic="true"
				data-fxs-announcer="polite"
				data-tabster-never-hide=""
				className={styles.hidden}
			>
				{polite.text}
			</div>
			<div
				aria-live="assertive"
				aria-atomic="true"
				data-fxs-announcer="assertive"
				data-tabster-never-hide=""
				className={styles.hidden}
			>
				{assertive.text}
			</div>
		</>
	);
	return (
		<AnnouncerContext.Provider value={announce}>
			<AnnounceProvider value={fluent}>
				{typeof document === "undefined" ? regions : createPortal(regions, document.body)}
				{children}
			</AnnounceProvider>
		</AnnouncerContext.Provider>
	);
}

const SILENT = { announce: () => {} };

/** Fluent components inside don't announce themselves (Properties' bars, #186) */
export function SilentAnnouncements({ children }: { children: ReactNode }) {
	return <AnnounceProvider value={SILENT}>{children}</AnnounceProvider>;
}

/**
 * A bar's text, to compare with what it announced: Fluent reads the body and the actions joined by a
 * comma ("Export failed. Disk full,Try again"), which the bar's own text doesn't have
 */
const comparable = (text: string) => text.replace(/[\s,]+/g, "");

/**
 * A tab's bars announce once per change (#186, owner decision): the tab panel unmounts when another
 * tab is chosen, so its bars would announce again on every visit ("The query didn't run", read
 * assertively each time you go back to Results). Inside a scope, a bar whose text was already read
 * isn't read again while it's still there or its tab is only hidden; once the bar goes away while
 * its tab is showing (a new run cleared the error), the text is forgotten and reads next time.
 */
export function AnnouncementScope({ shown, children }: { shown: boolean; children: ReactNode }) {
	const parent = useAnnounce();
	const heard = useRef(new Set<string>());
	const container = useRef<HTMLDivElement>(null);
	const latestParent = useRef(parent);
	useEffect(() => {
		latestParent.current = parent;
	});
	const value = useMemo(
		() => ({
			announce: (message: string, options?: { alert?: boolean; batchId?: string }) => {
				const text = message.trim();
				if (!text || heard.current.has(text)) return;
				heard.current.add(text);
				latestParent.current.announce(message, options);
			},
		}),
		[],
	);
	// Set during the commit, so a prune that runs after the tab was hidden (the bars went with it)
	// doesn't forget what was read
	const shownNow = useRef(shown);
	useLayoutEffect(() => {
		shownNow.current = shown;
	}, [shown]);
	// While the tab shows, forget texts whose bar has gone
	useEffect(() => {
		const el = container.current;
		if (!shown || !el) return;
		const prune = () => {
			if (!shownNow.current) return;
			const present = [...el.querySelectorAll(".fui-MessageBar")].map((b) =>
				comparable(b.textContent ?? ""),
			);
			for (const text of heard.current)
				if (!present.some((p) => p.includes(comparable(text)))) heard.current.delete(text);
		};
		const observer = new MutationObserver(prune);
		observer.observe(el, { childList: true, subtree: true, characterData: true });
		return () => observer.disconnect();
	}, [shown]);
	return (
		<div ref={container} style={{ display: "contents" }}>
			<AnnounceProvider value={value}>{children}</AnnounceProvider>
		</div>
	);
}
