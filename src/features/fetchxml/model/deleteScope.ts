/**
 * "Delete all" deletes what the query on screen matches (#174).
 *
 * "Delete all" (no selection) submits a Bulk Delete job for a query. It must be the query that
 * produced the visible results, not the tree after later edits: editing a filter in the left pane
 * doesn't re-run the query, so the grid keeps showing the old rows while the tree would match
 * different records.
 *
 * Edits that don't change which records match are fine: columns (attributes) don't, orders don't
 * unless `top` picks the first N, and an outer link-entity that only returns columns doesn't
 * restrict the root records. That covers what the results grid itself does (Edit columns, add or
 * remove columns, add related columns as outer joins, sort by a column header); its column filters
 * re-run the query.
 */
import { generateFetchXml } from "./fetchxml";
import type { EntityNode, FetchNode, FilterNode, LinkEntityNode } from "./nodes";

type Container = EntityNode | LinkEntityNode;

/** Whether a link-entity can change which root records match (anything but a column-only outer join) */
function restrictsRecords(link: LinkEntityNode): boolean {
	if (link.linkType !== "outer") return true;
	const hasFilter = (filters: FilterNode[] | undefined): boolean =>
		(filters ?? []).some((f) => f.conditions.length > 0 || (f.links?.length ?? 0) > 0 || hasFilter(f.subfilters));
	return hasFilter(link.filters) || link.links.some(restrictsRecords);
}

/**
 * The part of a query that decides which records match: the query without attributes, without
 * orders (unless `top` limits the rows), and without column-only outer link-entities. Two queries
 * with the same key match the same records.
 */
export function deleteScopeKey(fetch: FetchNode): string {
	const keepOrders = fetch.options?.top !== undefined;
	const strip = <T extends Container>(c: T): T => ({
		...c,
		attributes: [],
		allAttributes: undefined,
		orders: keepOrders ? c.orders : [],
		filters: c.filters.map(stripFilter),
		links: c.links.filter(restrictsRecords).map(strip),
	});
	const stripFilter = (f: FilterNode): FilterNode => ({
		...f,
		subfilters: f.subfilters.map(stripFilter),
		links: f.links?.map(strip),
	});
	return generateFetchXml({ ...fetch, entity: strip(fetch.entity) });
}

/**
 * Why "Delete all" isn't available, or undefined when it is: the results on screen must come from
 * a query that matches the same records as the tree.
 */
export function deleteAllUnavailableReason(executed: FetchNode | null | undefined, current: FetchNode | null | undefined): string | undefined {
	if (!current) return undefined;
	if (!executed) return "Run the query first: Delete all deletes the records the query on screen matches.";
	if (deleteScopeKey(executed) !== deleteScopeKey(current)) {
		return "The query has changed since these results were loaded. Run it again to see what Delete all would delete.";
	}
	return undefined;
}

/** Selected records up to this many are deleted by the tool itself; more go to a Bulk Delete job. */
export const MAX_DIRECT_DELETE = 100;

/**
 * Why "Delete" isn't available for the selected records, or undefined when it is (#177). More than
 * MAX_DIRECT_DELETE selected records are deleted by a Bulk Delete job, which needs prvBulkDelete.
 * Without it the tool used to fall back to deleting every one of them itself.
 */
export function deleteSelectionUnavailableReason(selectedCount: number, canBulkDelete: boolean): string | undefined {
	if (selectedCount <= MAX_DIRECT_DELETE || canBulkDelete) return undefined;
	return (
		`Deleting more than ${MAX_DIRECT_DELETE} records submits a Bulk Delete job, which needs the Bulk Delete ` +
		`privilege (prvBulkDelete). Select ${MAX_DIRECT_DELETE} or fewer, or ask your administrator for the privilege.`
	);
}
