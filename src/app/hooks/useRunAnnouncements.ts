/**
 * What a run does, for screen readers (#186 PR 3, owner decision 6): polite, through the app's
 * announcer, one batch so a newer status replaces a waiting one.
 *  - A run starts: "Running the query…" (as Run's reason says while it runs).
 *  - The first page arrives: "50 rows loaded, more available. Ran in 160 ms." (or "3 rows, all
 *    loaded.").
 *  - Retrieve all: progress at most every PROGRESS_EVERY_MS ("1,000 rows loaded so far."), not every
 *    page, then "5,000 rows, all loaded." when it's done.
 *  - A failed run, Stop, the Retrieve all limit and a failed page are MessageBars, which announce
 *    themselves (Announcer), so nothing is said here for them.
 * Pages loaded while scrolling aren't announced: the rows appear where the user is reading.
 */
import { useEffect, useRef } from "react";
import { useAnnouncer } from "../../shared/feedback/useAnnouncer";
import type { QueryExecution } from "./useQueryExecution";

export const PROGRESS_EVERY_MS = 10000;
const BATCH = "run-status";

const rowsText = (n: number) => `${n.toLocaleString()} ${n === 1 ? "row" : "rows"}`;

export function useRunAnnouncements(
	execution: Pick<
		QueryExecution,
		"isExecuting" | "isRetrieveAllInProgress" | "queryResult" | "executeError" | "pagingNotice"
	>,
): void {
	const announce = useAnnouncer();
	const { isExecuting, isRetrieveAllInProgress, queryResult, executeError, pagingNotice } =
		execution;
	const was = useRef({ isRetrieveAllInProgress });
	/** A run started and its first page (or its error) hasn't been told yet */
	const awaitingFirstPage = useRef(false);
	const lastProgress = useRef(0);

	// Start and first page
	useEffect(() => {
		if (isExecuting) {
			if (awaitingFirstPage.current) return;
			awaitingFirstPage.current = true;
			announce("Running the query…", { batchId: BATCH });
			return;
		}
		if (!awaitingFirstPage.current) return;
		// A failed run's bar says why (it announces itself); a run whose results were cleared (the
		// entity changed) has nothing to tell. Either way "Running…" isn't read late, after the error.
		if (executeError || !queryResult) {
			awaitingFirstPage.current = false;
			announce("", { batchId: BATCH });
			return;
		}
		awaitingFirstPage.current = false;
		const n = queryResult.rows.length;
		const count = isRetrieveAllInProgress
			? `${rowsText(n)} loaded, loading the rest`
			: queryResult.moreRecords
				? `${rowsText(n)} loaded, more available`
				: `${rowsText(n)}, all loaded`;
		const time =
			queryResult.executionTimeMs !== undefined
				? ` Ran in ${queryResult.executionTimeMs.toLocaleString()} ms.`
				: "";
		announce(`${count}.${time}`, { batchId: BATCH });
		lastProgress.current = Date.now();
	}, [isExecuting, isRetrieveAllInProgress, queryResult, executeError, announce]);

	// Retrieve all: progress, then done
	useEffect(() => {
		const before = was.current.isRetrieveAllInProgress;
		was.current.isRetrieveAllInProgress = isRetrieveAllInProgress;
		if (!queryResult || isExecuting) return;
		const n = queryResult.rows.length;
		if (isRetrieveAllInProgress) {
			if (Date.now() - lastProgress.current < PROGRESS_EVERY_MS) return;
			lastProgress.current = Date.now();
			announce(`${rowsText(n)} loaded so far.`, { batchId: BATCH });
			return;
		}
		// Done; Stop, the limit and a failed page say so in their own bar (the "no paging cookie"
		// note doesn't end anything)
		const ended = pagingNotice && pagingNotice.kind !== "simple";
		if (before && !ended) announce(`${rowsText(n)}, all loaded.`, { batchId: BATCH });
	}, [isRetrieveAllInProgress, queryResult, isExecuting, pagingNotice, announce]);
}
