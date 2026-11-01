/**
 * Paging helpers (#58, #60). Pure functions, verified against a live org (2026-10-16):
 *
 * - FetchXML through the Web API returns `@Microsoft.Dynamics.CRM.fetchxmlpagingcookie` (the
 *   `@Microsoft.Dynamics.CRM.pagingcookie` key FetchXML Studio used to read is never sent) and
 *   `@Microsoft.Dynamics.CRM.morerecords`. On the last page neither is sent.
 * - The annotation is an element: <cookie pagenumber="2" pagingcookie="…" istracking="False" />.
 *   `pagingcookie` is URL-encoded twice. `pagenumber` isn't used: with an empty cookie it stays at 2
 *   for every page (verified), so the next page is always the requested page + 1.
 * - Some queries (sorted by a link-entity attribute, distinct) return the annotation with an empty
 *   `pagingcookie`: they page with page + count only ("simple paging").
 * - Saved views run with ?savedQuery= / ?userQuery= return `@odata.nextLink` instead.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results
 */

export interface PagingCookieInfo {
	/** The decoded cookie for the next request's paging-cookie attribute, if Dataverse returned one */
	cookie?: string;
}

const attributeOf = (element: string, name: string): string | undefined => {
	try {
		const doc = new DOMParser().parseFromString(element, "application/xml");
		if (doc.querySelector("parsererror")) return undefined;
		return doc.documentElement.getAttribute(name) ?? undefined;
	} catch {
		return undefined;
	}
};

const decodeTwice = (value: string): string | undefined => {
	try {
		return decodeURIComponent(decodeURIComponent(value));
	} catch {
		return undefined;
	}
};

/**
 * Read the paging cookie from the `@Microsoft.Dynamics.CRM.fetchxmlpagingcookie` annotation: the
 * `pagingcookie` attribute, URL-decoded twice. Empty or missing means simple paging.
 */
export function parseFetchXmlPagingCookie(annotation: unknown): PagingCookieInfo {
	if (typeof annotation !== "string" || !annotation.trim()) return {};
	const encoded = attributeOf(annotation, "pagingcookie");
	const cookie = encoded ? decodeTwice(encoded) : undefined;
	return cookie ? { cookie } : {};
}

export interface FetchPagingOptions {
	top?: number;
	count?: number;
	page?: number;
}

/**
 * The <fetch> start tag. A quoted attribute value may contain a raw ">" (XML allows it), so quoted
 * values are skipped whole rather than ending the match early (#200).
 */
export const FETCH_START_TAG = /<fetch\b(?:[^>"']|"[^"]*"|'[^']*')*>/i;

/**
 * The paging attributes of the FetchXML that actually ran (not the tree's, which can differ),
 * from the <fetch> start tag. Attribute values may use either quote style.
 */
export function readFetchPagingOptions(fetchXml: string): FetchPagingOptions {
	const tag = FETCH_START_TAG.exec(fetchXml)?.[0] ?? "";
	const read = (name: string) => {
		const match = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, "i").exec(tag);
		const value = Number((match?.[1] ?? match?.[2] ?? "").trim());
		return match && Number.isInteger(value) ? value : undefined;
	};
	const options: FetchPagingOptions = {};
	const top = read("top");
	const count = read("count");
	const page = read("page");
	if (top !== undefined) options.top = top;
	if (count !== undefined) options.count = count;
	if (page !== undefined) options.page = page;
	return options;
}

/**
 * The path queryData expects for an @odata.nextLink (everything after /api/data/v9.x/); the host
 * prefixes the connection's URL and API version itself.
 */
export function nextLinkToQueryPath(nextLink: string): string | undefined {
	const match = /^https?:\/\/[^/]+\/api\/data\/v[\d.]+\/(.+)$/i.exec(nextLink.trim());
	return match?.[1];
}

/** The paging fields of one page's response (what pptbClient returns). */
export interface PageResponse {
	moreRecords?: boolean;
	pagingCookie?: string;
	nextLink?: string;
}

/** Paging position after the last loaded page (#58, #60). */
export interface PagingState {
	/** FetchXML pages with page + paging cookie; saved views follow @odata.nextLink */
	mode: "fetchxml" | "view";
	/** The FetchXML that ran (pages repeat it with paging attributes) */
	xml: string;
	/**
	 * The executed query's own count, kept for every page; smaller when a page was too large for the
	 * PPTB host's 16 MB response limit (#68, pageRecovery.ts)
	 */
	pageSize?: number;
	/**
	 * Rows in the first page when the query sets no count and has more pages: Dataverse's default
	 * page size for the Entity (5,000 for standard tables, 500 for elastic tables) (#68)
	 */
	fullPageSize?: number;
	currentPage: number;
	nextPage: number;
	pagingCookie?: string;
	nextLink?: string;
	moreRecords: boolean;
	isRetrieveAllInProgress: boolean;
	/** Pages are being retrieved with Retrieve All (Retry continues it) */
	retrieveAll: boolean;
}

/** Whether there's a page after `response`: saved views only page with @odata.nextLink. */
export function hasNextPage(mode: PagingState["mode"], response: PageResponse): boolean {
	return mode === "view" ? !!response.moreRecords && !!response.nextLink : !!response.moreRecords;
}

/**
 * The position after fetching page `state.nextPage`. The next page is always that page + 1:
 * Dataverse's `pagenumber` hint stays at 2 for queries without a paging cookie (verified), and
 * following it repeated page 2 forever.
 */
/**
 * After the first page of a query whose `count` is larger than the page Dataverse returned (#209):
 * the maximum page size is 5,000 rows (500 for elastic tables), and without a paging cookie later
 * pages are requested by page number, which skips (page − 1) × count rows. Requesting page 2 with
 * the query's count would then skip the rows between the returned page and `count`. So paging goes
 * on by the size Dataverse actually returned, from the row after the ones loaded. When that row
 * isn't the start of a page of that size (a starting page above 1), paging stops with a notice
 * rather than skip rows. With a paging cookie, or a full page, nothing changes.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results
 */
export function pagingAfterFirstPage(
	state: PagingState,
	firstPage: number,
	rowsReturned: number,
): { state: PagingState; notice?: PagingNotice } {
	const count = state.pageSize;
	if (state.mode !== "fetchxml" || state.pagingCookie || !state.moreRecords || !count || rowsReturned <= 0 || rowsReturned >= count) {
		return { state };
	}
	const loadedThrough = (firstPage - 1) * count + rowsReturned;
	if (loadedThrough % rowsReturned !== 0) {
		return {
			state: { ...state, moreRecords: false, isRetrieveAllInProgress: false },
			notice: {
				kind: "simple",
				message: `Dataverse returned ${rowsReturned.toLocaleString()} rows, fewer than the query's count of ${count.toLocaleString()}, and returned no paging cookie. Later pages can't be requested from page ${firstPage} without skipping rows, so they aren't loaded. Set count to ${rowsReturned.toLocaleString()} or less, or start from page 1.`,
			},
		};
	}
	const currentPage = loadedThrough / rowsReturned;
	return { state: { ...state, pageSize: rowsReturned, currentPage, nextPage: currentPage + 1 } };
}

export function advancePaging(state: PagingState, response: PageResponse): PagingState {
	return {
		...state,
		currentPage: state.nextPage,
		nextPage: state.nextPage + 1,
		pagingCookie: response.pagingCookie,
		nextLink: response.nextLink,
		moreRecords: hasNextPage(state.mode, response),
	};
}

/** What the results pane tells the user about paging (#60). */
export type PagingNotice =
	| { kind: "error"; page?: number; message: string }
	| { kind: "stopped"; message: string }
	| { kind: "limit"; message: string }
	| { kind: "simple"; message: string };

/**
 * Retrieve All stops loading once this many rows are loaded (#183): every row is kept in the
 * webview's memory, with its formatted values, so an unbounded run on a wide table can exhaust it.
 * Further pages load one at a time by scrolling, as after Stop.
 */
export const RETRIEVE_ALL_ROW_LIMIT = 100_000;

/** Whether Retrieve All stops before fetching another page, with this many rows loaded */
export const retrieveAllLimitReached = (
	loadedRows: number,
	limit = RETRIEVE_ALL_ROW_LIMIT,
): boolean => loadedRows >= limit;

/** The notice when Retrieve All stops at the limit */
export const RETRIEVE_ALL_LIMIT_MESSAGE =
	"Retrieve all stops once 100,000 rows or more are loaded, to keep the tool responsive. Scroll to the end of the grid to load more one page at a time. To delete every matching record without loading them, use Bulk delete (job).";
