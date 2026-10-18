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
	[2048, "Template"],
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
