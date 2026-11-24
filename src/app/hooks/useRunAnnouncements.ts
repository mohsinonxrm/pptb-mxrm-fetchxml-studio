/**
 * What a run does, for screen readers (#186 PR 3, owner decision 6): polite, through the app's
 * announcer, one batch so a newer status replaces a waiting one.
 *  - A run starts: "Running query…".
 *  - The first page arrives: "50 rows loaded, more available. Ran in 160 ms." (or "all loaded").
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
			announce("Running query…", { batchId: BATCH });
			return;
		}
		if (!awaitingFirstPage.current) return;
		// A failed run's bar says why (it announces itself)
		if (executeError) {
			awaitingFirstPage.current = false;
			return;
		}
		if (!queryResult) return;
		awaitingFirstPage.current = false;
		const n = queryResult.rows.length;
		const more = isRetrieveAllInProgress
			? ", loading the rest"
			: queryResult.moreRecords
				? ", more available"
				: ", all loaded";
		const time =
			queryResult.executionTimeMs !== undefined
				? ` Ran in ${queryResult.executionTimeMs.toLocaleString()} ms.`
				: "";
		announce(`${rowsText(n)} loaded${more}.${time}`, { batchId: BATCH });
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
		// Done; Stop, the limit and a failed page say so in their own bar
		if (before && !pagingNotice) announce(`${rowsText(n)}, all loaded.`, { batchId: BATCH });
	}, [isRetrieveAllInProgress, queryResult, isExecuting, pagingNotice, announce]);
}
