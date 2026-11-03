/**
 * Saved view helpers (#93): system views (savedquery) and personal views (userquery).
 * Pure functions only, so the load / execute / save rules can be unit tested.
 */

import { getParseToTreeChanges, type ParseResult } from "./fetchxmlParser";

/**
 * Whether the query about to run is the loaded view, unmodified, so it can run by its ID
 * (?savedQuery= / ?userQuery=), which pages with @odata.nextLink.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/retrieve-and-execute-predefined-queries
 *
 * Both sides must be generated from a tree (generateFetchXml): the view's stored FetchXML carries
 * version / mapping / descending="false" and other defaults the generator leaves out, so comparing
 * the stored text with generated text never matched (#93).
 */
export function isUnmodifiedView(currentFetchXml: string, viewFetchXml: string): boolean {
	return currentFetchXml.trim() === viewFetchXml.trim();
}

/**
 * Query types of savedquery / userquery records (the `querytype` column).
 *  - The five view types you can customize: Public 0, Advanced Find 1, Associated 2, Quick Find 4,
 *    Lookup 64. https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#types-of-views
 *  - The same column also stores Outlook / offline filters and templates and other queries:
 *    https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.savedqueryquerytype
 * userquery uses the same values ("address book filter, advanced search, or other view"):
 * https://learn.microsoft.com/power-apps/developer/data-platform/reference/entities/userquery#BKMK_QueryType
 */
export const PUBLIC_VIEW_QUERY_TYPE = 0;

const QUERY_TYPE_LABELS: ReadonlyArray<readonly [number, string]> = [
	[1, "Advanced find"],
	[2, "Associated"],
	[4, "Quick find"],
	[64, "Lookup"],
	[8, "Reporting"],
	[16, "Offline filter"],
	[128, "Appointment book"],
	[256, "Outlook filter"],
	[512, "Address book filter"],
	[1024, "Without subject"],
	[2048, "Other"], // SavedQueryTypeOther (#144)
	[4096, "Dialog"],
	[8192, "Offline template"],
	[16384, "Custom"],
	[65536, "Export translations"],
	[131072, "Outlook template"],
];

/** Display order of the query types: public views first, then the order above, unknown types last. */
const QUERY_TYPE_ORDER = new Map<number, number>([
	[PUBLIC_VIEW_QUERY_TYPE, 0],
	...QUERY_TYPE_LABELS.map(([type], index): [number, number] => [type, index + 1]),
]);

/**
 * Short label shown next to a view that isn't a public view, so every view of the Entity is listed
 * and told apart rather than hidden (#93). Undefined for public views.
 */
export function getQueryTypeLabel(queryType: number | undefined): string | undefined {
	if (queryType === undefined || queryType === PUBLIC_VIEW_QUERY_TYPE) return undefined;
	const known = QUERY_TYPE_LABELS.find(([type]) => type === queryType);
	return known ? known[1] : `Query type ${queryType}`;
}

/** A view's owner (personal views): the ownerid lookup and its formatted value (#142) */
interface OwnedView {
	type: "system" | "personal";
	ownerId?: string;
	ownerName?: string;
}

/**
 * Whether a personal view belongs to someone else: a user query "can be assigned and shared with
 * other users" (https://learn.microsoft.com/power-apps/developer/data-platform/saved-queries), and
 * Load view lists the ones shared with the user (#93). False when either ID is unknown (#142).
 */
export function isSharedView(view: OwnedView, currentUserId: string | undefined): boolean {
	if (view.type !== "personal" || !view.ownerId || !currentUserId) return false;
	return view.ownerId.toLowerCase() !== currentUserId.toLowerCase();
}

/**
 * Second line of a Load view option (#142): "Shared by {owner}" for a personal view someone else
 * owns, the query type for views that aren't public views ("Quick find view"), "Default view" for
 * the default public view. Undefined for the user's own public views.
 */
export function getViewSecondaryText(
	view: OwnedView & { queryType?: number; isDefault?: boolean },
	currentUserId: string | undefined,
): string | undefined {
	const parts: string[] = [];
	if (isSharedView(view, currentUserId)) parts.push(view.ownerName ? `Shared by ${view.ownerName}` : "Shared with you");
	const typeLabel = getQueryTypeLabel(view.queryType);
	if (typeLabel) {
		const known = QUERY_TYPE_LABELS.some(([, label]) => label === typeLabel);
		parts.push(known ? `${typeLabel} view` : typeLabel);
	} else if (view.isDefault && view.type === "system") {
		parts.push("Default view");
	}
	return parts.length > 0 ? parts.join(" · ") : undefined;
}

/**
 * Whether RetrievePrincipalAccess's AccessRights include WriteAccess (#142). The Web API returns the
 * AccessRights flags as a comma-separated string ("ReadAccess, WriteAccess, ..."; verified live);
 * a number is read as the flags value (WriteAccess = 2). Undefined when it can't be read.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/retrieveprincipalaccess
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/accessrights
 */
export function hasWriteAccess(accessRights: unknown): boolean | undefined {
	if (typeof accessRights === "number") return (accessRights & 2) === 2;
	if (typeof accessRights !== "string") return undefined;
	return accessRights
		.split(",")
		.map((right) => right.trim())
		.includes("WriteAccess");
}

/**
 * Whether "Set as default view" applies (#142): public system views only. "You can set only one
 * active public view as the default view. To make a view the default view, set the IsDefault
 * property to true." Personal views (userquery) have no isdefault column.
 * https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#set-a-public-view-as-the-default-view
 */
export function canSetDefaultView(view: { type: "system" | "personal"; queryType?: number }): boolean {
	return view.type === "system" && (view.queryType ?? PUBLIC_VIEW_QUERY_TYPE) === PUBLIC_VIEW_QUERY_TYPE;
}

/**
 * Quick find placeholder (#142). Quick find views store `value="{0}"` in the conditions of their
 * quick find filter, and the caller puts the search string in its place ("string.Format" in the
 * Learn samples): https://learn.microsoft.com/power-apps/developer/data-platform/quick-find
 * Dataverse doesn't substitute it: a view with {0} run as stored matches nothing (verified live).
 */
/**
 * The warning after a view became the default public view while another view stayed a default too:
 * Dataverse doesn't clear it, and clearing it failed (#307). An entry without a name means the other
 * defaults couldn't be read.
 */
export function defaultViewsNotClearedMessage(viewName: string, notCleared: readonly { name: string; error: string }[]): string {
	const named = notCleared.filter((view) => view.name);
	const errors = notCleared.map((view) => view.error).join(" ");
	if (named.length === 0) {
		return `"${viewName}" is the default view now, but FetchXML Studio couldn't check whether another view is still marked as the default too. To be sure, set the default again in Power Apps (the entity's Views, then Set as default). Details: ${errors}`;
	}
	const views = named.map((view) => `"${view.name}"`).join(", ");
	return `"${viewName}" is the default view now, but ${views} ${named.length === 1 ? "is" : "are"} still marked as the default too. Set the default again in Power Apps (the entity's Views, then Set as default) to fix it. Details: ${errors}`;
}

export const QUICK_FIND_PLACEHOLDER = "{0}";

/** Condition values: the value attribute of <condition> (either quote) and <value> elements. */
const CONDITION_VALUE_PATTERNS: RegExp[] = [
	/(<condition\b[^>]*?\svalue=")([^"]*)(")/g,
	/(<condition\b[^>]*?\svalue=')([^']*)(')/g,
	/(<value>)([^<]*)(<\/value>)/g,
];

/** The number of condition values holding the quick find placeholder {0} (#142). */
export function countQuickFindPlaceholders(fetchXml: string): number {
	let count = 0;
	for (const pattern of CONDITION_VALUE_PATTERNS) {
		for (const match of fetchXml.matchAll(pattern)) {
			if (match[2].includes(QUICK_FIND_PLACEHOLDER)) count++;
		}
	}
	return count;
}

/**
 * The search value a quick find test value runs with (#142): quick find conditions use like, which
 * "requires a search string that ends with %", and the Learn samples append it when it's missing.
 * https://learn.microsoft.com/power-apps/developer/data-platform/quick-find#what-is-a-quick-find-query
 */
export function toQuickFindSearchValue(value: string, addTrailingPercent: boolean): string {
	const trimmed = value.trim();
	if (!trimmed) return "";
	return addTrailingPercent && !trimmed.endsWith("%") ? `${trimmed}%` : trimmed;
}

/**
 * A copy of the FetchXML with {0} replaced by the search value in every condition value (#142),
 * XML-escaped for the attribute or element it goes into. For execution only: the tree, the editor
 * and saved views keep {0}.
 */
export function applyQuickFindValue(fetchXml: string, searchValue: string): string {
	const escape = (quote: string) => (text: string) => {
		let escaped = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
		if (quote === '"') escaped = escaped.replace(/"/g, "&quot;");
		if (quote === "'") escaped = escaped.replace(/'/g, "&apos;");
		return escaped;
	};
	return CONDITION_VALUE_PATTERNS.reduce(
		(xml, pattern) =>
			xml.replace(pattern, (whole, open: string, value: string, close: string) =>
				value.includes(QUICK_FIND_PLACEHOLDER)
					? `${open}${value.split(QUICK_FIND_PLACEHOLDER).join(escape(close === "</value>" ? "" : close)(searchValue))}${close}`
					: whole,
			),
		fetchXml,
	);
}

interface SortableView {
	name: string;
	queryType?: number;
	isDefault?: boolean;
}

/** Public views first (the default view on top), then the other query types; by name within each. */
export function sortViewsByQueryType<T extends SortableView>(views: T[]): T[] {
	const rank = (view: T) =>
		QUERY_TYPE_ORDER.get(view.queryType ?? PUBLIC_VIEW_QUERY_TYPE) ?? QUERY_TYPE_ORDER.size;
	return [...views].sort(
		(a, b) =>
			rank(a) - rank(b) ||
			(a.queryType ?? 0) - (b.queryType ?? 0) ||
			Number(b.isDefault === true) - Number(a.isDefault === true) ||
			a.name.localeCompare(b.name),
	);
}

/**
 * Overall result of ValidateFetchXmlExpression, which returns "warnings and suggestions about the
 * performance of the query", each with a Severity from 0 (low) to 3 (critical). High and critical
 * show as errors; it is advice only and never blocks saving (#93).
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/validatefetchxmlexpression
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/validatorissue
 */
export function getValidationStatus(
	messages: ReadonlyArray<{ Severity: number }>,
): "valid" | "warning" | "error" {
	if (messages.length === 0) return "valid";
	return messages.some((m) => m.Severity >= 2) ? "error" : "warning";
}

/**
 * Whether a loaded view can be overwritten (#93). A system view is updated only "if the
 * IsCustomizable managed property allows the view to be updated":
 * https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#update-views
 * Personal views have no such property, and an unknown value isn't blocked (Dataverse decides).
 */
export function canOverwriteView(view: { type: "system" | "personal"; isCustomizable?: boolean }): boolean {
	return view.type !== "system" || view.isCustomizable !== false;
}

/**
 * Reads a ManagedProperty column (e.g. savedquery.iscustomizable) from a Web API record: the Web API
 * returns a BooleanManagedProperty complex type, whose Value is the setting.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/booleanmanagedproperty
 */
export function readManagedPropertyValue(value: unknown): boolean | undefined {
	if (typeof value === "boolean") return value;
	if (value && typeof value === "object" && typeof (value as { Value?: unknown }).Value === "boolean") {
		return (value as { Value: boolean }).Value;
	}
	return undefined;
}

export interface ViewLoadNotice {
	title: string;
	body: string;
	type: "warning" | "error";
}

/**
 * What to tell the user after loading a view's FetchXML into the tree (#93): nothing when it loaded
 * as-is, an error when it couldn't be parsed, a warning listing what the tree changed (saving the
 * view would save those changes).
 */
export function getViewLoadNotice(viewName: string, result: ParseResult): ViewLoadNotice | null {
	if (!result.success) {
		const reason = result.errors[0]?.message;
		return {
			title: "Could not load view",
			body: `The FetchXML of "${viewName}" couldn't be parsed${reason ? `: ${reason}` : "."}`,
			type: "error",
		};
	}
	const changes = getParseToTreeChanges(result);
	if (changes.length === 0) return null;
	const list = changes.map((c) => (c.element ? `<${c.element}> ${c.message}` : c.message)).join(" ");
	return {
		title: "View loaded with changes",
		body: `The tree can't represent everything in "${viewName}". ${list} Saving the view saves these changes.`,
		type: "warning",
	};
}
