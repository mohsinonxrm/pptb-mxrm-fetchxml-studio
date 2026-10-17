/**
 * Column filters from the results grid (#23)
 *
 * "Filter by…" on a column header adds a <condition> to the query, so the filter covers every
 * record, not just the loaded ones. The conditions live in their own <filter type="and"> group in
 * the root entity's filter, apart from the conditions built in the tree:
 *
 * - An entity has at most one <filter> (entity element reference), so the group is nested in it.
 *   With no root filter, the group becomes the root filter.
 * - A root filter of type "or" is wrapped in a new "and" root filter, so the column filters narrow
 *   the results instead of widening them. The query means the same otherwise. The wrapper's ID is
 *   kept, so only that wrapper is unwrapped later, never an "and" filter the user wrote.
 * - Link-entity columns use entityname, which filters after the join, so outer joins filter the
 *   rows shown (Filters on link-entity).
 *
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/entity
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows#filters-on-link-entity
 */

import type { ConditionNode, FetchNode, FilterNode, NodeId } from "./nodes";
import { resolveColumn } from "./queryColumns";

/** What a column filter's condition targets */
export interface ColumnFilterTarget {
	attribute: string;
	/** For link-entity columns: the link-entity's alias, or its name when it has none */
	entityname?: string;
	/** Logical name of the entity that holds the attribute (for its metadata) */
	entityName: string;
}

/** The condition a column filter targets, or why the column can't be filtered */
export function columnFilterTarget(fetch: FetchNode, columnKey: string): ColumnFilterTarget | { reason: string } {
	if (fetch.options?.aggregate) {
		return { reason: "Column filters aren't available in aggregate queries. Add conditions in the tree instead." };
	}
	const column = resolveColumn(fetch, columnKey);
	if (!column) return { reason: "This column isn't an attribute in the query." };
	if (column.isRoot) return { attribute: column.attribute.name, entityName: fetch.entity.name };
	const link = column.container;
	return {
		attribute: column.attribute.name,
		entityname: "alias" in link && link.alias ? link.alias : link.name,
		entityName: link.name,
	};
}

export type NewCondition = Omit<ConditionNode, "id" | "type">;

/** Path to a filter: index in the root entity's filters, then subfilter indexes */
type FilterPath = number[];

function findPath(filters: FilterNode[], id: NodeId, path: FilterPath = []): FilterPath | undefined {
	for (const [i, filter] of filters.entries()) {
		if (filter.id === id) return [...path, i];
		const inner = findPath(filter.subfilters, id, [...path, i]);
		if (inner) return inner;
	}
	return undefined;
}

function filterAt(filters: FilterNode[], path: FilterPath): FilterNode | undefined {
	let list = filters;
	let found: FilterNode | undefined;
	for (const i of path) {
		found = list[i];
		if (!found) return undefined;
		list = found.subfilters;
	}
	return found;
}

/** Replaces the filter at a path (undefined removes it) */
function replaceAt(filters: FilterNode[], path: FilterPath, replace: (f: FilterNode) => FilterNode | undefined): FilterNode[] {
	const [i, ...rest] = path;
	return filters.flatMap((filter, index) => {
		if (index !== i) return [filter];
		if (rest.length === 0) {
			const next = replace(filter);
			return next ? [next] : [];
		}
		return [{ ...filter, subfilters: replaceAt(filter.subfilters, rest, replace) }];
	});
}

const withRootFilters = (fetch: FetchNode, filters: FilterNode[]): FetchNode => ({
	...fetch,
	entity: { ...fetch.entity, filters },
});

/** The column-filter group, if it's still in the query */
export function findColumnFilterGroup(fetch: FetchNode | null, groupId: NodeId | null): FilterNode | undefined {
	if (!fetch || !groupId) return undefined;
	const path = findPath(fetch.entity.filters, groupId);
	return path ? filterAt(fetch.entity.filters, path) : undefined;
}

/** The column filters' conditions (none when the group was removed or never added) */
export function columnFilterConditions(fetch: FetchNode | null, groupId: NodeId | null): ConditionNode[] {
	return findColumnFilterGroup(fetch, groupId)?.conditions ?? [];
}

/** Adds a column filter's condition, creating the group when needed */
export function addColumnFilter(
	fetch: FetchNode,
	groupId: NodeId | null,
	condition: NewCondition,
	generateId: () => NodeId,
): { fetch: FetchNode; groupId: NodeId; conditionId: NodeId; wrapperId?: NodeId } {
	const newCondition: ConditionNode = { ...condition, id: generateId(), type: "condition" };
	const filters = fetch.entity.filters;
	const path = groupId ? findPath(filters, groupId) : undefined;
	if (groupId && path) {
		return {
			fetch: withRootFilters(fetch, replaceAt(filters, path, (g) => ({ ...g, conditions: [...g.conditions, newCondition] }))),
			groupId,
			conditionId: newCondition.id,
		};
	}

	const group: FilterNode = {
		id: generateId(),
		type: "filter",
		conjunction: "and",
		conditions: [newCondition],
		subfilters: [],
		links: [],
	};
	const result = (next: FilterNode[]) => ({ fetch: withRootFilters(fetch, next), groupId: group.id, conditionId: newCondition.id });

	const [root, ...others] = filters;
	// No root filter: the group is the root filter
	if (!root) return result([group]);
	// An "and" root filter: nest the group in it
	if (root.conjunction !== "or") return result([{ ...root, subfilters: [...root.subfilters, group] }, ...others]);
	// An "or" root filter: wrap it, so the column filters still narrow the results
	const wrapper: FilterNode = {
		id: generateId(),
		type: "filter",
		conjunction: "and",
		conditions: [],
		subfilters: [root, group],
		links: [],
	};
	return { ...result([wrapper, ...others]), wrapperId: wrapper.id };
}

/** Removes a column filter's condition; the group goes when it's empty */
export function removeColumnFilter(
	fetch: FetchNode,
	groupId: NodeId | null,
	conditionId: NodeId,
	wrapperId: NodeId | null = null,
): { fetch: FetchNode; groupId: NodeId | null } {
	const group = findColumnFilterGroup(fetch, groupId);
	if (!group || !groupId) return { fetch, groupId: null };
	const conditions = group.conditions.filter((c) => c.id !== conditionId);
	if (conditions.length > 0 || group.subfilters.length > 0 || group.links.length > 0) {
		const path = findPath(fetch.entity.filters, groupId)!;
		return { fetch: withRootFilters(fetch, replaceAt(fetch.entity.filters, path, (g) => ({ ...g, conditions }))), groupId };
	}
	return { fetch: clearColumnFilters(fetch, groupId, wrapperId), groupId: null };
}

/** Removes the column-filter group, and the "and" wrapper addColumnFilter put around an "or" root filter */
export function clearColumnFilters(fetch: FetchNode, groupId: NodeId | null, wrapperId: NodeId | null = null): FetchNode {
	const path = groupId ? findPath(fetch.entity.filters, groupId) : undefined;
	if (!path) return fetch;
	const filters = replaceAt(fetch.entity.filters, path, () => undefined);
	// Undo the wrapping of an "or" root filter: only the wrapper addColumnFilter made, and only while
	// it still holds just that filter (the user may have added to it since)
	const [root, ...others] = filters;
	if (
		path.length === 2 &&
		root &&
		root.id === wrapperId &&
		root.conjunction === "and" &&
		root.conditions.length === 0 &&
		root.links.length === 0 &&
		root.subfilters.length === 1 &&
		!root.hint &&
		!root.isQuickFindFields &&
		!root.overrideQuickFindRecordLimitEnabled &&
		!root.overrideQuickFindRecordLimitDisabled &&
		!root.extraAttributes
	) {
		return withRootFilters(fetch, [root.subfilters[0], ...others]);
	}
	return withRootFilters(fetch, filters);
}

const conditionsKey = (filter: FilterNode | undefined) =>
	JSON.stringify(
		filter?.conditions.map((c) => [c.attribute, c.operator, c.value ?? null, c.entityname ?? null]) ?? null,
	);

/** The wrapper's ID after the query was loaded again: the new root filter, when the group was relocated */
export function relocateColumnFilterWrapper(
	before: FetchNode | null,
	after: FetchNode,
	wrapperId: NodeId | null,
	relocatedGroupId: NodeId | null,
): NodeId | null {
	if (!before || !wrapperId || !relocatedGroupId) return null;
	return before.entity.filters[0]?.id === wrapperId ? (after.entity.filters[0]?.id ?? null) : null;
}

/**
 * The group's ID after the query was loaded again (editor sync gives every node a new ID): the
 * filter at the same place, when it has the same conditions.
 */
export function relocateColumnFilterGroup(before: FetchNode | null, after: FetchNode, groupId: NodeId | null): NodeId | null {
	if (!before || !groupId) return null;
	const path = findPath(before.entity.filters, groupId);
	if (!path) return null;
	const previous = filterAt(before.entity.filters, path);
	const candidate = filterAt(after.entity.filters, path);
	return candidate && conditionsKey(candidate) === conditionsKey(previous) ? candidate.id : null;
}
