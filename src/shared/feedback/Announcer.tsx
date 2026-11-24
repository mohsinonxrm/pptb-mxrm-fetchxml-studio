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
 *   (Properties, whose bars would re-announce on every node selection).
 *
 * Fluent's own AriaLiveAnnouncer isn't used: without `ariaNotify` its region is always assertive.
 *
 * Messages that arrive close together are read as one (FLUSH_MS), in order; a message with the same
 * `batchId` as a waiting one replaces it, so a burst ("Undone" pressed five times, Retrieve all
 * progress) reads its last step. The text is cleared before each new reading, so the same message
 * twice is read twice, and again after CLEAR_AFTER_MS, so a later read of the page doesn't meet it.
 */
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { AnnounceProvider, makeStyles } from "@fluentui/react-components";
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
			setText(joined);
			timers.current.clear = setTimeout(() => setText(""), CLEAR_AFTER_MS);
		}, FLUSH_MS);
	}, []);
	return { text, add };
}

export function AnnouncerProvider({ children }: { children: ReactNode }) {
	const styles = useStyles();
	const polite = useRegion();
	const assertive = useRegion();
	const { add: addPolite } = polite;
	const { add: addAssertive } = assertive;
	const announce = useCallback<Announce>(
		(message, options) => {
			const text = message.trim();
			if (!text) return;
			(options?.politeness === "assertive" ? addAssertive : addPolite)(
				text,
				options?.batchId,
			);
		},
		[addPolite, addAssertive],
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
	return (
		<AnnouncerContext.Provider value={announce}>
			<AnnounceProvider value={fluent}>
				<div
					role="status"
					aria-live="polite"
					aria-atomic="true"
					data-fxs-announcer="polite"
					className={styles.hidden}
				>
					{polite.text}
				</div>
				<div
					aria-live="assertive"
					aria-atomic="true"
					data-fxs-announcer="assertive"
					className={styles.hidden}
				>
					{assertive.text}
				</div>
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
