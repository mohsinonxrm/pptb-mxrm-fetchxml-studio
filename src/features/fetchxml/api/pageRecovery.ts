/**
 * Fetching the next page despite throttling and the PPTB host's 16 MB response limit (#68, #96).
 *
 * PPTB 1.2.7 rejects any Dataverse response over 16 MB with "Dataverse response exceeds the maximum
 * size" (dataverseManager.ts makeHttpRequest, MAX_DATAVERSE_BATCH_RESPONSE_BYTES). A 5,000-row page
 * of a wide Entity with all annotations can exceed it (verified on the sandbox: Learn's
 * all-attributes account example). When a page fails that way, the same page is requested again
 * with half the count, and paging continues at that size.
 * https://github.com/PowerPlatformToolBox/desktop-app/blob/v1.2.7/src/main/managers/dataverseManager.ts
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results
 *
 * Only FetchXML paging can change its page size; saved views follow @odata.nextLink, whose page
 * size Dataverse sets.
 */

import { isResponseTooLargeError } from "./dataverseErrors";
import type { PagingState } from "./paging";
import { withThrottleRetry, type ThrottleRetryOptions } from "./throttleRetry";

/**
 * Dataverse's default (and maximum) page size for standard tables, the starting point for halving a
 * first page that sets no count (#183).
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results
 */
export const DATAVERSE_DEFAULT_PAGE_SIZE = 5000;

/** No smaller page than this is tried (a page that still doesn't fit fails with the friendly error). */
export const MIN_ADAPTIVE_PAGE_SIZE = 100;
/** At most this many halvings for one page (5,000 → 2,500 → 1,250 → 625). */
export const MAX_PAGE_HALVINGS = 3;

/**
 * The paging position for requesting the same next page with half the page size, or undefined
 * when it can't shrink (saved view, unknown or odd size, below the minimum).
 *
 * - With a paging cookie, Dataverse continues after the cookie's last record, so the page number
 *   stays and only count changes.
 * - Without one (simple paging), Dataverse skips (page - 1) × count rows: at half the count, the
 *   same first row is page 2 × (page - 1) + 1. Halving only even sizes keeps that exact.
 * Neither has been verified live yet (#68 needs-verification).
 */
export function halvePageSize(state: PagingState): PagingState | undefined {
	if (state.mode !== "fetchxml") return undefined;
	const size = state.pageSize ?? state.fullPageSize;
	if (!size || !Number.isInteger(size) || size % 2 !== 0) return undefined;
	const half = size / 2;
	if (half < MIN_ADAPTIVE_PAGE_SIZE) return undefined;
	const nextPage = state.pagingCookie ? state.nextPage : 2 * (state.nextPage - 1) + 1;
	return { ...state, pageSize: half, nextPage };
}

export interface PageRecoveryOptions {
	/** Back-off after service protection errors (#96) */
	retry?: ThrottleRetryOptions;
	maxHalvings?: number;
	/** Called when the page size is halved (debug logging) */
	onShrink?: (state: PagingState) => void;
}

/**
 * Fetch the page after `state` with `fetchPage`: retried after service protection errors, and
 * with half the page size after a response-too-large error (bounded). Returns the result and the
 * paging position it was fetched with (its pageSize is what later pages use).
 */
export async function fetchPageWithRecovery<T>(
	state: PagingState,
	fetchPage: (state: PagingState) => Promise<T>,
	{ retry, maxHalvings = MAX_PAGE_HALVINGS, onShrink }: PageRecoveryOptions = {},
): Promise<{ result: T; state: PagingState }> {
	let current = state;
	for (let halvings = 0; ; halvings++) {
		try {
			const result = await withThrottleRetry(() => fetchPage(current), retry);
			return { result, state: current };
		} catch (error) {
			const smaller =
				halvings < maxHalvings && isResponseTooLargeError(error)
					? halvePageSize(current)
					: undefined;
			if (!smaller) throw error;
			onShrink?.(smaller);
			current = smaller;
		}
	}
}
