/**
 * Running the query and paging its results (#301): Execute and the first page, Load more, Retrieve
 * all with Stop and its 100,000-row limit (#60, #183), the paging notice and its Retry, the throttle
 * wait while a Retrieve all page waits to retry (#96, #183), and the execute error.
 *
 * Results belong to one run (#60): a newer run, Stop, clearing the results (the entity changes, a
 * view or New query) or unmounting drops whatever is still loading for an older one. AppShell
 * remounts this hook's component when the tool's connection changes (#66), so the unmount is what
 * stops a Retrieve all from the previous connection. Retrieve all holds the host's close warning
 * while it runs (#70) and releases it however it ends.
 *
 * AppShell calls the hook where the results were cleared on an entity change, so that effect keeps
 * its place among AppShell's effects.
 */
import { useEffect, useLayoutEffect, useRef, useState } from "react";
import {
	executeFetchXml,
	executeNextLink,
	executeSystemView,
	executePersonalView,
} from "../../features/fetchxml/api/pptbClient";
import type { FetchXmlResult } from "../../features/fetchxml/api/pptbClient";
import { addPagingToFetchXml } from "../../features/fetchxml/model/fetchxml";
import {
	advancePaging,
	hasNextPage,
	readFetchPagingOptions,
	retrieveAllLimitReached,
	RETRIEVE_ALL_LIMIT_MESSAGE,
	type PagingNotice,
	type PagingState,
} from "../../features/fetchxml/api/paging";
import { formatDataverseError } from "../../features/fetchxml/api/dataverseErrors";
import {
	DATAVERSE_DEFAULT_PAGE_SIZE,
	fetchPageWithRecovery,
} from "../../features/fetchxml/api/pageRecovery";
import {
	LONG_RUNNING_RETRY,
	withThrottleRetry,
	type ThrottleRetryOptions,
} from "../../features/fetchxml/api/throttleRetry";
import { holdClose } from "../../features/fetchxml/api/host";
import {
	applyQuickFindValue,
	countQuickFindPlaceholders,
	isUnmodifiedView,
} from "../../features/fetchxml/model/savedViews";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import type { LoadedViewState } from "../../features/fetchxml/state/builderStore";
import type { QueryResult } from "../../features/fetchxml/ui/RightPane/ResultsGrid";
import { debugLog } from "../../shared/utils/debug";

export const SIMPLE_PAGING_MESSAGE =
	"Dataverse didn't return a paging cookie for this query (for example, it's sorted by a link-entity attribute or uses distinct), so pages are fetched with page numbers only. Microsoft recommends paging cookies; this works, but is slower for large result sets.";

/** In editor mode, the editor's FetchXML and the tree it was loaded into, run straight away (#102) */
export interface SyncedQuery {
	xml: string;
	fetchQuery: FetchNode;
}

export interface QueryExecutionInputs {
	/** The builder's query; its root <entity> node decides when the results are cleared */
	fetchQuery: FetchNode | null;
	/** The tree's FetchXML, as AppShell generates it (nothing injected, #59) */
	fetchXml: string;
	/** The quick find test value that replaces {0} in the FetchXML that runs (#142); "" for none */
	quickFindSearchValue: string;
	/** The view the query was loaded from or saved as; run by its ID while it's unmodified (#93) */
	loadedView: LoadedViewState | null;
	/** The query's "Retrieve all records" option (#60) */
	retrieveAllRecords: boolean;
}

export interface QueryExecution {
	/** The results in the grid, or null before a run and after they're cleared */
	queryResult: QueryResult | null;
	/** The first page is loading */
	isExecuting: boolean;
	/** A further page is loading, or Retrieve all is running */
	isLoadingMore: boolean;
	isRetrieveAllInProgress: boolean;
	/** The current paging state; its xml is the FetchXML that ran (useRecordActions, #142) */
	pagingRef: { readonly current: PagingState | null };
	/** A failed page, Stop, the Retrieve all limit, or simple paging (#60, #183) */
	pagingNotice: PagingNotice | null;
	/** While a Retrieve all page waits to retry because Dataverse is throttling: the wait (#183) */
	pagingWaitSeconds: number | undefined;
	/** Why the last run failed, shown above the results (#96) */
	executeError: string | undefined;

	/** Runs the tree's query, or the synced editor query (#102), and Retrieve all when it's on */
	execute: (synced?: SyncedQuery) => Promise<void>;
	/** The latest execute, for callbacks that run the query again after an await (#65) */
	handleExecuteRef: { readonly current: (synced?: SyncedQuery) => Promise<void> };
	/** The next page, for infinite scroll when Retrieve all is off */
	loadMore: () => Promise<void>;
	stopRetrieveAll: () => void;
	/** Retries the page that failed: Retrieve all continues, or the one page loads (#60) */
	retryPaging: () => void;
	dismissExecuteError: () => void;
	dismissPagingNotice: () => void;
}

export function useQueryExecution({
	fetchQuery,
	fetchXml,
	quickFindSearchValue,
	loadedView: builderLoadedView,
	retrieveAllRecords,
}: QueryExecutionInputs): QueryExecution {
	// State for query execution
	const [queryResult, setQueryResult] = useState<QueryResult | null>(null);
	const [isExecuting, setIsExecuting] = useState(false);
	const [isLoadingMore, setIsLoadingMore] = useState(false);

	// Paging for infinite scroll / Retrieve All (#58, #60). The ref is what the async page loops read
	// (no stale closures); the state mirrors it for rendering.
	const pagingRef = useRef<PagingState | null>(null);
	const [pagingState, setPagingStateValue] = useState<PagingState | null>(null);
	const setPaging = (next: PagingState | null) => {
		pagingRef.current = next;
		setPagingStateValue(next);
	};
	// What happened to paging that the user should know about: a failed page, Stop, simple paging
	const [pagingNotice, setPagingNotice] = useState<PagingNotice | null>(null);
	// Run tokens (#60): results from an older execution, or pages from a stopped loop, are dropped.
	// resultRun changes with every execution and when the results are cleared (entity change, view,
	// New); pageRun also changes on Stop, which keeps the loaded rows.
	const resultRunRef = useRef(0);
	const pageRunRef = useRef(0);
	// Rows loaded for the current results, for Retrieve All's limit (#183)
	const loadedRowsRef = useRef(0);
	// While a Retrieve All page waits to retry because Dataverse is throttling: the wait (#183)
	const [pagingWaitSeconds, setPagingWaitSeconds] = useState<number | undefined>(undefined);
	// Unmounting (e.g. the connection changed, #66) stops page loops still running for these results
	useEffect(
		() => () => {
			resultRunRef.current++;
			pageRunRef.current++;
		},
		[],
	);

	// Execution error (e.g. Dataverse API error surfaced to user)
	const [executeError, setExecuteError] = useState<string | undefined>();

	// Entity node of the tree the current results were produced for (see the clearing effect below)
	const resultEntityIdRef = useRef<string | undefined>(undefined);

	// Clear query results when the fetch query structure changes (entity change, view clear, etc.)
	// We use the entity node id as a proxy - it changes when entity is re-selected or view is cleared.
	// Results produced for the new tree are kept: an editor-mode Execute loads the editor's FetchXML
	// into the tree and runs it straight away, and a fast response can arrive first (#102).
	useEffect(() => {
		const entityId = fetchQuery?.entity?.id;
		if (entityId !== undefined && entityId === resultEntityIdRef.current) return;
		// Pages still loading for the old results must not bring them back (#60)
		resultRunRef.current++;
		pageRunRef.current++;
		setQueryResult(null);
		setPaging(null);
		setPagingNotice(null);
		setIsLoadingMore(false);
		setPagingWaitSeconds(undefined);
		// A first page still loading for the old results is dropped, so it can't end the spinner
		setIsExecuting(false);
	}, [fetchQuery?.entity?.id]);

	/**
	 * Execute the tree's query and handle Retrieve All if enabled.
	 * @param synced - in editor mode the editor's FetchXML is loaded into the tree first (#102);
	 *   the new tree and its XML are passed here because builder state updates on the next render.
	 */
	const handleExecute = async (synced?: { xml: string; fetchQuery: FetchNode }) => {
		const query = synced?.fetchQuery ?? fetchQuery;
		const builtXml = synced?.xml ?? fetchXml;
		if (!builtXml) return;
		// Quick find views (#142): the test value replaces {0} in the FetchXML that runs (and its later
		// pages); the tree, the editor and saved views keep {0}. Without a value nothing runs: the
		// test value field asks for one.
		// https://learn.microsoft.com/power-apps/developer/data-platform/quick-find
		const hasQuickFindPlaceholders = countQuickFindPlaceholders(builtXml) > 0;
		if (hasQuickFindPlaceholders && !quickFindSearchValue) return;
		const xmlToExecute = hasQuickFindPlaceholders
			? applyQuickFindValue(builtXml, quickFindSearchValue)
			: builtXml;

		// A new run: anything still loading for the previous one is dropped (#60)
		const run = ++resultRunRef.current;
		pageRunRef.current++;

		setIsExecuting(true);
		setIsLoadingMore(false);
		setPagingWaitSeconds(undefined);
		setQueryResult(null);
		setPaging(null);
		setPagingNotice(null);
		setExecuteError(undefined);

		resultEntityIdRef.current = query?.entity.id;

		const startTime = performance.now();
		const entityLogicalName = query?.entity.name;
		const retrieveAll = retrieveAllRecords;

		// Paging follows the FetchXML that runs, not the tree (#60). Microsoft: don't page a query that
		// uses top; Dataverse applies the limit.
		const options = readFetchPagingOptions(xmlToExecute);
		const hasTopLimit = options.top !== undefined;

		try {
			// Determine execution method based on loaded view state. A synced editor query
			// replaced the tree (and cleared the loaded view), so it always runs as FetchXML.
			let result: FetchXmlResult | undefined;
			// Set when the first page needed a smaller count (#183); later pages keep it
			let firstPageShrunkTo: PagingState | undefined;
			const loadedView = synced ? null : builderLoadedView;
			let useViewExecution = false;

			// A quick find view runs as FetchXML with the test value: by ID it would run with {0} as
			// stored, which matches nothing (verified live, #142)
			if (loadedView && !hasQuickFindPlaceholders) {
				if (isUnmodifiedView(xmlToExecute, loadedView.originalFetchXml)) {
					useViewExecution = true;
					debugLog(
						"app",
						`📋 Executing ${loadedView.type} view "${loadedView.name}" via ${
							loadedView.type === "system" ? "savedQuery" : "userQuery"
						}=${loadedView.id}`,
					);

					// Retried after service protection errors like every page (#183, #96); a view's page
					// size is Dataverse's, so it can't be made smaller
					if (loadedView.type === "system") {
						result = await withThrottleRetry(() =>
							executeSystemView(loadedView.entitySetName, loadedView.id),
						);
					} else {
						result = await withThrottleRetry(() =>
							executePersonalView(loadedView.entitySetName, loadedView.id),
						);
					}
				}
			}

			// If not using view execution (either no view, view was modified, or editor override), use FetchXML
			if (!result) {
				debugLog(
					"app",
					synced
						? "✏️ Executing editor FetchXML (loaded into the tree)"
						: hasQuickFindPlaceholders
							? "🔎 Executing quick find FetchXML with the test value"
							: loadedView
								? `📝 View "${loadedView.name}" was modified - executing its FetchXML`
								: "📡 Executing FetchXML query",
				);
				// The first page gets the same recovery as later pages (#183): retried after service
				// protection errors (#96), and requested again with half the count when it's over the
				// PPTB host's 16 MB response limit (#68). That is the common case: a 5,000-row first page
				// of a wide Entity. Without a count, Dataverse's default page size (5,000) is the starting
				// point; queries with top or aggregate can't take a count, so they only get the retry.
				const canShrink = !hasTopLimit && query?.options?.aggregate !== true;
				const firstPageState: PagingState = {
					mode: "fetchxml",
					xml: xmlToExecute,
					pageSize: canShrink
						? (options.count ?? DATAVERSE_DEFAULT_PAGE_SIZE)
						: undefined,
					currentPage: (options.page ?? 1) - 1,
					nextPage: options.page ?? 1,
					moreRecords: true,
					isRetrieveAllInProgress: false,
					retrieveAll: false,
				};
				const { result: firstPageResult, state: fetchedFirst } =
					await fetchPageWithRecovery(
						firstPageState,
						// Unchanged, the query runs as written; after a halving, with the smaller count
						(s) =>
							s.pageSize === firstPageState.pageSize
								? executeFetchXml(xmlToExecute)
								: executeFetchXml(
										addPagingToFetchXml(
											xmlToExecute,
											s.nextPage,
											undefined,
											s.pageSize,
										),
									),
						{
							onShrink: (s) =>
								debugLog(
									"app",
									`First page too large: retrying with count ${s.pageSize}`,
								),
						},
					);
				result = firstPageResult;
				if (fetchedFirst.pageSize !== firstPageState.pageSize)
					firstPageShrunkTo = fetchedFirst;
			}
			if (run !== resultRunRef.current) return; // superseded while the first page was loading

			const executionTimeMs = Math.round(performance.now() - startTime);
			const moreRecords =
				!hasTopLimit && hasNextPage(useViewExecution ? "view" : "fetchxml", result);
			const firstPage = options.page ?? 1;

			loadedRowsRef.current = result.records.length;
			setQueryResult({
				query,
				rows: result.records.map((record) => ({ ...record })),
				totalRecordCount: result.totalRecordCount,
				totalRecordCountLimitExceeded: result.totalRecordCountLimitExceeded,
				moreRecords,
				pagingCookie: result.pagingCookie,
				entityLogicalName,
				executionTimeMs,
			});
			setPaging({
				mode: useViewExecution ? "view" : "fetchxml",
				xml: xmlToExecute,
				pageSize: firstPageShrunkTo?.pageSize ?? options.count,
				// Dataverse's default page size, for a smaller page after the 16 MB limit (#68)
				fullPageSize:
					options.count === undefined && moreRecords ? result.records.length : undefined,
				// After a smaller first page (#183), paging continues from the page it was fetched as
				currentPage: firstPageShrunkTo?.nextPage ?? firstPage,
				nextPage: (firstPageShrunkTo?.nextPage ?? firstPage) + 1,
				pagingCookie: result.pagingCookie,
				nextLink: result.nextLink,
				moreRecords,
				isRetrieveAllInProgress: retrieveAll && moreRecords,
				retrieveAll,
			});
			setIsExecuting(false);

			if (retrieveAll && moreRecords) await loadAllPages(run);
		} catch (error) {
			if (run !== resultRunRef.current) return;
			console.error("Failed to execute FetchXML:", error);
			// Documented Dataverse errors get actionable text; Dataverse's own message follows (#96)
			setExecuteError(
				formatDataverseError(
					error,
					"An unexpected error occurred while executing the query.",
				),
			);
			setQueryResult({ query, rows: [] });
			setIsExecuting(false);
		}
	};

	// The latest handleExecute, for callbacks that re-run the query after an await (record deletes):
	// they call the current version instead of listing a function that changes every render (#65)
	const handleExecuteRef = useRef(handleExecute);
	useLayoutEffect(() => {
		handleExecuteRef.current = handleExecute;
	});

	/**
	 * Fetch the page after `state` (#58, #60): saved views follow @odata.nextLink; FetchXML sets page,
	 * the decoded paging cookie (when Dataverse returned one) and the executed query's own count.
	 */
	const fetchNextPage = async (
		state: PagingState,
		retry?: ThrottleRetryOptions,
	): Promise<{ result: FetchXmlResult; next: PagingState }> => {
		// Retried after service protection errors (#96); a page over the host's 16 MB limit is
		// requested again with half the count, which later pages keep (#68)
		const { result, state: fetched } = await fetchPageWithRecovery(
			state,
			(s) =>
				s.mode === "view"
					? executeNextLink(s.nextLink ?? "")
					: executeFetchXml(
							addPagingToFetchXml(s.xml, s.nextPage, s.pagingCookie, s.pageSize),
						),
			{
				retry,
				onShrink: (s) =>
					debugLog(
						"app",
						`Page too large: retrying page ${s.nextPage} with count ${s.pageSize}`,
					),
			},
		);
		return { result, next: advancePaging(fetched, result) };
	};

	/** Add a fetched page to the results of run `run` (ignored if the results changed meanwhile). */
	const appendPage = (
		run: number,
		state: PagingState,
		result: FetchXmlResult,
		next: PagingState,
	) => {
		if (run !== resultRunRef.current) return false;
		loadedRowsRef.current += result.records.length;
		setQueryResult((prev) =>
			prev
				? {
						...prev,
						rows: prev.rows.concat(result.records),
						moreRecords: next.moreRecords,
						pagingCookie: next.pagingCookie,
					}
				: prev,
		);
		// A page loaded, so an earlier failure or Stop notice no longer applies. A FetchXML page
		// requested without a cookie means Dataverse returned none for this query: say so once.
		const simple: PagingNotice | null =
			state.mode === "fetchxml" && !state.pagingCookie
				? { kind: "simple", message: SIMPLE_PAGING_MESSAGE }
				: null;
		setPagingNotice((prev) =>
			prev && prev.kind !== "error" && prev.kind !== "stopped" && prev.kind !== "limit"
				? prev
				: simple,
		);
		return true;
	};

	/**
	 * Retrieve All: fetch the remaining pages one after another (#60). Stops when there are no more
	 * pages, when Stop is pressed (pageRun changes), when the results change (resultRun changes) or
	 * at RETRIEVE_ALL_ROW_LIMIT rows (#183). Throttled pages get the long back-off, with the wait in
	 * the status line (#183).
	 */
	const loadAllPages = async (run: number) => {
		const pageRun = ++pageRunRef.current;
		const isCurrent = () => run === resultRunRef.current && pageRun === pageRunRef.current;
		let state = pagingRef.current;
		if (!state) return;
		state = { ...state, isRetrieveAllInProgress: true, retrieveAll: true };
		setIsLoadingMore(true);
		setPaging(state);
		// The host warns before the tab closes while Retrieve All runs (#70)
		const releaseClose = holdClose();
		const retry: ThrottleRetryOptions = {
			...LONG_RUNNING_RETRY,
			// Stop or new results during a wait: the page isn't fetched again
			shouldContinue: isCurrent,
			onRetry: ({ delayMs }) => {
				if (isCurrent()) setPagingWaitSeconds(Math.ceil(delayMs / 1000));
			},
		};
		try {
			while (state?.moreRecords) {
				if (retrieveAllLimitReached(loadedRowsRef.current)) {
					// Further pages load one at a time from here, as after Stop
					state = { ...state, isRetrieveAllInProgress: false, retrieveAll: false };
					setPaging(state);
					setPagingNotice({ kind: "limit", message: RETRIEVE_ALL_LIMIT_MESSAGE });
					debugLog(
						"app",
						`Retrieve All stopped at ${loadedRowsRef.current} rows (limit)`,
					);
					return;
				}
				const { result, next } = await fetchNextPage(state, retry);
				if (!isCurrent()) return;
				setPagingWaitSeconds(undefined);
				appendPage(run, state, result, next);
				state = { ...next, isRetrieveAllInProgress: next.moreRecords };
				setPaging(state);
			}
			debugLog("app", "✅ Retrieve All complete");
		} catch (error) {
			if (!isCurrent()) return;
			console.error("Failed during Retrieve All:", error);
			setPagingNotice({
				kind: "error",
				page: state?.nextPage,
				message: formatDataverseError(error), // (#96)
			});
		} finally {
			releaseClose();
			if (isCurrent()) {
				setPagingWaitSeconds(undefined);
				setIsLoadingMore(false);
				if (pagingRef.current)
					setPaging({ ...pagingRef.current, isRetrieveAllInProgress: false });
			}
		}
	};

	/** Stop Retrieve All (#60): the page in flight is discarded; the loaded rows stay. */
	const handleStopRetrieveAll = () => {
		const state = pagingRef.current;
		if (!state?.isRetrieveAllInProgress) return;
		pageRunRef.current++;
		setIsLoadingMore(false);
		setPagingWaitSeconds(undefined);
		// Further pages load one at a time from here (scrolling, or Retry after a failed page)
		setPaging({ ...state, isRetrieveAllInProgress: false, retrieveAll: false });
		setPagingNotice({ kind: "stopped", message: "" });
	};

	/**
	 * Load more records (for infinite scroll when Retrieve All is OFF)
	 */
	const handleLoadMore = async () => {
		const state = pagingRef.current;
		if (!state || !state.moreRecords || state.isRetrieveAllInProgress || isLoadingMore) return;

		const run = resultRunRef.current;
		const pageRun = ++pageRunRef.current;
		const isCurrent = () => run === resultRunRef.current && pageRun === pageRunRef.current;
		setIsLoadingMore(true);

		try {
			const { result, next } = await fetchNextPage(state);
			if (!isCurrent()) return;
			appendPage(run, state, result, next);
			setPaging(next);
			debugLog("app", `✅ Page ${next.currentPage} loaded: ${result.records.length} records`);
		} catch (error) {
			if (!isCurrent()) return;
			console.error("Failed to load more records:", error);
			setPagingNotice({
				kind: "error",
				page: state.nextPage,
				message: formatDataverseError(error), // (#96)
			});
		} finally {
			if (isCurrent()) setIsLoadingMore(false);
		}
	};

	/** Retry the page that failed: continue Retrieve All, or load the one page (#60). */
	const handleRetryPaging = () => {
		setPagingNotice(null);
		const state = pagingRef.current;
		if (!state) return;
		if (state.retrieveAll) void loadAllPages(resultRunRef.current);
		else void handleLoadMore();
	};

	return {
		queryResult,
		isExecuting,
		isLoadingMore,
		isRetrieveAllInProgress: pagingState?.isRetrieveAllInProgress ?? false,
		pagingRef,
		pagingNotice,
		pagingWaitSeconds,
		executeError,
		execute: handleExecute,
		handleExecuteRef,
		loadMore: handleLoadMore,
		stopRetrieveAll: handleStopRetrieveAll,
		retryPaging: handleRetryPaging,
		dismissExecuteError: () => setExecuteError(undefined),
		dismissPagingNotice: () => setPagingNotice(null),
	};
}
