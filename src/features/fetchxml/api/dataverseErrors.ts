/**
 * Dataverse errors made actionable (#96, #68). Pure: no host or React dependency.
 *
 * How errors reach FetchXML Studio (PPTB host v1.2.7, src/main/managers/dataverseManager.ts):
 * - A Dataverse error response rejects with `new Error(`${error.code}: ${error.message}`)`, from the
 *   OData error JSON, e.g. "0x8004e023: AggregateQueryRecordLimit exceeded. Cannot perform this
 *   operation." The HTTP status and the response headers (including Retry-After) are not passed on.
 * - A response without an OData error body rejects with "HTTP <status>" (and the body, if any).
 * - A response over the host's 16 MB cap rejects with "Dataverse response exceeds the maximum size"
 *   (metadata requests: "Dataverse metadata response exceeds the maximum size") (#68).
 * - Electron's IPC layer may prefix "Error invoking remote method '<channel>': Error: ".
 * https://github.com/PowerPlatformToolBox/desktop-app/blob/v1.2.7/src/main/managers/dataverseManager.ts
 *
 * Only codes documented on Microsoft Learn are mapped; any other error is shown exactly as it
 * arrived. The mapped text never replaces Dataverse's own message: it's shown before it.
 * https://learn.microsoft.com/power-apps/developer/data-platform/reference/web-service-error-codes
 */

/** What an error is about, for callers that react to it (retry, smaller page). */
export type DataverseErrorKind =
	| "serviceProtection"
	| "responseTooLarge"
	| "queryThrottling"
	| "queryLimit"
	| "query"
	| "privilege"
	| "notFound"
	| "unknown";

export interface DataverseErrorInfo {
	/** The Dataverse hex error code as sent (e.g. "0x8004e023"), when there is one */
	code?: string;
	/** The HTTP status, when the host passed it on ("HTTP 429") or the error object carries it */
	status?: number;
	/** The error text as received, without Electron's IPC prefix: Dataverse's own message stays */
	message: string;
	/** What the user can do about it, for documented errors */
	guidance?: string;
	/** The Microsoft Learn (or PPTB) page that documents the error */
	docUrl?: string;
	kind: DataverseErrorKind;
	/** Seconds to wait before retrying, when the error carries a Retry-After value */
	retryAfterSeconds?: number;
}

interface KnownError {
	kind: DataverseErrorKind;
	guidance: string;
	docUrl: string;
}

const LEARN = "https://learn.microsoft.com/power-apps/developer/data-platform";
const ERROR_CODES_URL = `${LEARN}/reference/web-service-error-codes`;
const API_LIMITS_URL = `${LEARN}/api-limits`;
const ANTIPATTERNS_URL = `${LEARN}/query-antipatterns`;

const SERVICE_PROTECTION_GUIDANCE =
	"Dataverse service protection limits were reached. Wait a few minutes before running more requests, or work with fewer records at a time.";

/**
 * Documented codes, keyed by the lower-case hex code. Each entry cites the Learn article that
 * documents the code and the fix.
 */
const KNOWN_CODES: Record<string, KnownError> = {
	// Service protection API limits: Web API returns 429 with a Retry-After header.
	// https://learn.microsoft.com/power-apps/developer/data-platform/api-limits#service-protection-api-limit-errors
	"0x80072322": {
		kind: "serviceProtection",
		guidance: SERVICE_PROTECTION_GUIDANCE,
		docUrl: `${API_LIMITS_URL}#number-of-requests`,
	},
	"0x80072321": {
		kind: "serviceProtection",
		guidance: SERVICE_PROTECTION_GUIDANCE,
		docUrl: `${API_LIMITS_URL}#execution-time`,
	},
	"0x80072326": {
		kind: "serviceProtection",
		guidance: SERVICE_PROTECTION_GUIDANCE,
		docUrl: `${API_LIMITS_URL}#concurrent-requests`,
	},

	// Aggregate query record limit (50,000). Add filters and run in parts, or use aggregatelimit.
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data#aggregate-query-record-limit-error
	"0x8004e023": {
		kind: "queryLimit",
		guidance:
			"The aggregate query covers more than 50,000 records. Add filters (for example a date range) and run the query in parts, or set aggregatelimit on fetch to aggregate fewer records.",
		docUrl: `${LEARN}/fetchxml/aggregate-data#aggregate-query-record-limit-error`,
	},
	// QuickFindQueryRecordLimitExceeded
	// https://learn.microsoft.com/power-apps/developer/data-platform/quick-find#quick-find-record-limits
	"0x8004e024": {
		kind: "queryLimit",
		guidance:
			"The quick find search matches more records than the quick find record limit. Refine the search value or add filters.",
		docUrl: `${LEARN}/quick-find#quick-find-record-limits`,
	},
	// TooManyConditionsInQuery: up to 500 condition and link-entity elements in total.
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows#condition-limits
	"0x8004430c": {
		kind: "queryLimit",
		guidance:
			"The query has more than 500 condition and link-entity elements in total. Reduce them, for example by combining values into one condition with the in operator.",
		docUrl: `${LEARN}/fetchxml/filter-rows#condition-limits`,
	},
	// TooManyLinkEntitiesInQuery: up to 15 link-entity elements.
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables#limitations
	"0x8004430d": {
		kind: "queryLimit",
		guidance: "The query has more than 15 link-entity elements. Remove some link-entities.",
		docUrl: `${LEARN}/fetchxml/join-tables#limitations`,
	},

	// Query throttling. https://learn.microsoft.com/power-apps/developer/data-platform/query-throttling
	"0x80048544": {
		kind: "queryThrottling",
		guidance:
			"Dataverse throttled this query because it affects database health. Simplify the query (fewer attributes, link-entities and expensive conditions), then try again.",
		docUrl: `${LEARN}/query-throttling`,
	},
	"0x80048644": {
		kind: "queryThrottling",
		guidance:
			"Dataverse throttled this query because a condition uses a leading wildcard (a value that starts with %). Remove the leading wildcard, for example by using begins-with.",
		docUrl: `${ANTIPATTERNS_URL}#PerformanceLeadingWildCard`,
	},
	"0x80048744": {
		kind: "queryThrottling",
		guidance:
			"Dataverse throttled this query because a condition filters on a calculated or formula attribute. Remove that condition.",
		docUrl: `${ANTIPATTERNS_URL}#FilteringOnCalculatedColumns`,
	},
	// The message lists the anti-patterns found, comma separated.
	"0x80048745": {
		kind: "queryThrottling",
		guidance:
			"Dataverse throttled this query because of the query anti-patterns named in the message. Fix them as described in Query anti-patterns.",
		docUrl: `${LEARN}/query-throttling#dataverse-error-for-query-throttling-caused-by-anti-patterns`,
	},
	// Timeouts with a detected anti-pattern. https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns
	"0x80048573": {
		kind: "queryThrottling",
		guidance:
			"The query timed out, possibly because a condition uses a leading wildcard. Remove the leading wildcard, for example by using begins-with.",
		docUrl: `${ANTIPATTERNS_URL}#PerformanceLeadingWildCard`,
	},
	"0x80048574": {
		kind: "queryThrottling",
		guidance:
			"The query timed out, possibly because a condition filters on a calculated or formula attribute. Remove that condition.",
		docUrl: `${ANTIPATTERNS_URL}#FilteringOnCalculatedColumns`,
	},
	"0x80048575": {
		kind: "queryThrottling",
		guidance:
			"The query timed out because of the performance issues named in the message. Fix them as described in Query anti-patterns.",
		docUrl: ANTIPATTERNS_URL,
	},

	// Aggregates. https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data
	// QueryBuilderAttributeNotFound ("A required attribute was not specified"): seen live for an
	// attribute without aggregate or groupby in an aggregate query (#96).
	"0x8004111e": {
		kind: "query",
		guidance:
			"A required attribute is missing. In an aggregate query, every attribute needs an aggregate function or groupby, and an alias.",
		docUrl: `${LEARN}/fetchxml/aggregate-data`,
	},
	// QueryBuilderDeserializeInvalidAggregate: seen live for an aggregate attribute in a query
	// without aggregate="true" (#96).
	"0x8004111a": {
		kind: "query",
		guidance:
			'Dataverse couldn\'t process the aggregates in this query. Aggregate attributes need aggregate="true" on fetch, and each one needs an alias.',
		docUrl: `${LEARN}/fetchxml/aggregate-data`,
	},
	// QueryBuilderInvalidAttributeValue
	"0x80041139": {
		kind: "query",
		guidance:
			"A condition value isn't valid for its operator or attribute (for example, the last-x-days operators need a whole number).",
		docUrl: `${LEARN}/fetchxml/reference/operators`,
	},
	// QueryBuilderNoAttribute
	"0x80041103": {
		kind: "query",
		guidance:
			"An attribute in the query doesn't exist on its Entity. Check the attribute's logical name named in the message.",
		docUrl: ERROR_CODES_URL,
	},
	// QueryBuilderNoEntity
	"0x80041102": {
		kind: "query",
		guidance: "An Entity in the query wasn't found. Check the Entity's logical name.",
		docUrl: ERROR_CODES_URL,
	},
	// CrmHttpClientError: e.g. "Resource not found for the segment '…'" (an unknown entity set name).
	// https://learn.microsoft.com/troubleshoot/power-platform/dataverse/dataverse-web-api-and-sdk/web-api-client-errors#resource-not-found-for-the-segment
	"0x80060888": {
		kind: "notFound",
		guidance:
			"The Web API didn't recognize part of the request, such as an Entity's entity set name.",
		docUrl: "https://learn.microsoft.com/troubleshoot/power-platform/dataverse/dataverse-web-api-and-sdk/web-api-client-errors#resource-not-found-for-the-segment",
	},
	// ObjectDoesNotExist
	"0x80040217": {
		kind: "notFound",
		guidance: "The record or object wasn't found. It might have been deleted.",
		docUrl: ERROR_CODES_URL,
	},

	// PrivilegeDenied: the message names the missing privilege and the Entity.
	// https://learn.microsoft.com/power-platform/admin/security-roles-privileges
	"0x80040220": {
		kind: "privilege",
		guidance:
			"You're missing the privilege named in the message. Ask your administrator to add it to one of your security roles.",
		docUrl: "https://learn.microsoft.com/power-platform/admin/security-roles-privileges",
	},
};

/** HTTP 429 without an OData error body: the same service protection limits (api-limits). */
const TOO_MANY_REQUESTS: KnownError = {
	kind: "serviceProtection",
	guidance: SERVICE_PROTECTION_GUIDANCE,
	docUrl: API_LIMITS_URL,
};

/**
 * The PPTB host's 16 MB response cap (v1.2.7, MAX_DATAVERSE_BATCH_RESPONSE_BYTES) (#68).
 * https://github.com/PowerPlatformToolBox/desktop-app/blob/v1.2.7/src/main/utilities/dataverseBatch.ts#L10
 */
const RESPONSE_TOO_LARGE: KnownError = {
	kind: "responseTooLarge",
	guidance:
		"The response is larger than the 16 MB limit of Power Platform ToolBox. Set a smaller page size (count on fetch), select fewer attributes, or avoid all-attributes.",
	docUrl: "https://github.com/PowerPlatformToolBox/desktop-app/blob/v1.2.7/src/main/managers/dataverseManager.ts",
};

const IPC_PREFIX = /^Error invoking remote method '[^']*':\s*(?:Error:\s*)?/i;
const HEX_CODE = /\b0x[0-9a-f]{8}\b/i;
const HTTP_STATUS = /\bHTTP (\d{3})\b/;

/**
 * Whether an error is the host's "response too large" rejection (#68). The host rejects with exactly
 * "Dataverse response exceeds the maximum size" (or "Dataverse metadata response exceeds the
 * maximum size" for $metadata), see dataverseManager.ts makeHttpRequest in PPTB v1.2.7.
 */
const RESPONSE_TOO_LARGE_TEXT = /Dataverse (?:metadata )?response exceeds the maximum size/i;

const readString = (value: unknown): string | undefined =>
	typeof value === "string" && value.trim() ? value : undefined;

const readNumber = (value: unknown): number | undefined => {
	const n = typeof value === "number" ? value : typeof value === "string" ? Number(value) : NaN;
	return Number.isFinite(n) ? n : undefined;
};

/**
 * Retry-After in seconds, if the error object carries one. The v1.2.7 host doesn't pass headers on;
 * this reads them should a later host (or the OData error) include them.
 */
function readRetryAfter(error: Record<string, unknown>): number | undefined {
	const headers = (error.headers ??
		(error.response as Record<string, unknown> | undefined)?.headers) as
		Record<string, unknown> | undefined;
	const value = error.retryAfter ?? headers?.["retry-after"] ?? headers?.["Retry-After"];
	const seconds = readNumber(value);
	return seconds !== undefined && seconds >= 0 ? seconds : undefined;
}

/** The text of any thrown value: Error, string, OData error JSON ({ error: { code, message } }). */
function messageOf(error: unknown): string | undefined {
	if (typeof error === "string") return readString(error);
	if (error instanceof Error) return readString(error.message);
	if (error && typeof error === "object") {
		const record = error as Record<string, unknown>;
		const odata = record.error as Record<string, unknown> | undefined;
		if (odata && typeof odata === "object") {
			const code = readString(odata.code);
			const message = readString(odata.message);
			if (code && message) return `${code}: ${message}`;
			if (message) return message;
		}
		return readString(record.message);
	}
	return undefined;
}

/**
 * Parse any error from the host or Dataverse (#96). Never loses the original text: `message` is
 * what arrived, minus Electron's IPC prefix.
 */
export function parseDataverseError(
	error: unknown,
	fallback = "An unexpected error occurred.",
): DataverseErrorInfo {
	const raw = messageOf(error);
	const message = raw ? raw.replace(IPC_PREFIX, "").trim() || raw : fallback;
	const record = error && typeof error === "object" ? (error as Record<string, unknown>) : {};
	const odataCode = readString((record.error as Record<string, unknown> | undefined)?.code);
	const code =
		(odataCode && HEX_CODE.test(odataCode) ? odataCode : undefined) ??
		HEX_CODE.exec(message)?.[0];
	const status =
		readNumber(record.status ?? record.statusCode) ??
		readNumber(HTTP_STATUS.exec(message)?.[1]);
	const retryAfterSeconds = readRetryAfter(record);

	const known = RESPONSE_TOO_LARGE_TEXT.test(message)
		? RESPONSE_TOO_LARGE
		: ((code ? KNOWN_CODES[code.toLowerCase()] : undefined) ??
			(status === 429 ? TOO_MANY_REQUESTS : undefined));

	const info: DataverseErrorInfo = { message, kind: known?.kind ?? "unknown" };
	if (code) info.code = code;
	if (status !== undefined) info.status = status;
	if (known) {
		info.guidance = known.guidance;
		info.docUrl = known.docUrl;
	}
	if (retryAfterSeconds !== undefined) info.retryAfterSeconds = retryAfterSeconds;
	return info;
}

/**
 * The text to show for an error (#96): the guidance for a documented error, then the original
 * message after "Details:"; anything else exactly as it arrived.
 */
export function formatDataverseError(error: unknown, fallback?: string): string {
	const info = parseDataverseError(error, fallback);
	return info.guidance ? `${info.guidance} Details: ${info.message}` : info.message;
}

/** A service protection limit error (429 / 0x80072322 / 0x80072321 / 0x80072326): safe to retry later. */
export function isServiceProtectionError(error: unknown): boolean {
	return parseDataverseError(error).kind === "serviceProtection";
}

/** The PPTB host's 16 MB response cap was hit (#68). */
export function isResponseTooLargeError(error: unknown): boolean {
	return parseDataverseError(error).kind === "responseTooLarge";
}
