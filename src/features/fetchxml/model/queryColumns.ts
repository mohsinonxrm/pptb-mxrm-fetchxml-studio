/**
 * Result columns of a query and how they map back to the tree (#114). One model shared by the
 * layout, the results grid, Edit columns (remove) and header sorting, so they can't disagree.
 *
 * Column keys are the property names Dataverse returns, verified live (2026-10-15):
 * - an attribute with an alias: the alias;
 * - a root entity attribute: its name (lookups also come back as _name_value);
 * - a link-entity attribute: "linkalias.name"; a link-entity without an alias gets "{name}{N}",
 *   where N is its 1-based position among the joined link-entities in document order (aliased ones
 *   count, link-entities inside filters don't). Documented on the link-entity reference (alias).
 */

import type {
	AttributeNode,
	EntityNode,
	FetchNode,
	FilterNode,
	LinkEntityNode,
	OrderNode,
} from "./nodes";
import { findOrdersUsingAlias } from "./treeUtils";
import { getOrderMetadataIssue, type AttributeTypeInfo } from "./attributeCapabilities";

export interface QueryColumn {
	/** Property name in the results */
	key: string;
	attribute: AttributeNode;
	/** Entity / link-entity that holds the attribute */
	container: EntityNode | LinkEntityNode;
	isRoot: boolean;
	/** For link-entity attributes: the alias Dataverse uses for the link-entity ("pc", "contact1") */
	linkAlias?: string;
}

/** The alias Dataverse uses for each joined link-entity (its alias, or "{name}{N}"). */
export function joinedLinkAliases(fetch: FetchNode): Map<string, string> {
	const result = new Map<string, string>();
	let position = 0;
	const visit = (container: EntityNode | LinkEntityNode) =>
		container.links.forEach((link) => {
			position++;
			result.set(link.id, link.alias || `${link.name}${position}`);
			visit(link);
		});
	visit(fetch.entity);
	return result;
}

/** Every attribute the query returns, with its result column key, in document order. */
export function collectQueryColumns(fetch: FetchNode): QueryColumn[] {
	const aliases = joinedLinkAliases(fetch);
	const columns: QueryColumn[] = fetch.entity.attributes.map((attribute) => ({
		key: attribute.alias || attribute.name,
		attribute,
		container: fetch.entity,
		isRoot: true,
	}));
	const visit = (container: EntityNode | LinkEntityNode) =>
		container.links.forEach((link) => {
			const linkAlias = aliases.get(link.id)!;
			for (const attribute of link.attributes) {
				columns.push({
					key: attribute.alias || `${linkAlias}.${attribute.name}`,
					attribute,
					container: link,
					isRoot: false,
					linkAlias,
				});
			}
			visit(link);
		});
	visit(fetch.entity);
	return columns;
}

/**
 * The query column behind a result/layout column key. Also accepts the _name_value lookup form, and the
 * schema name a matchfirstrowusingcrossapply link-entity returns its attributes by (#91): the logical
 * name is the schema name in lowercase.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables#related-table-property-and-attribute-names-are-inconsistent
 */
export function resolveColumn(fetch: FetchNode, key: string): QueryColumn | undefined {
	const columns = collectQueryColumns(fetch);
	const direct = columns.find((c) => c.key === key);
	if (direct) return direct;
	const lookup = /^_(.+)_value$/.exec(key)?.[1];
	const root = lookup
		? columns.find((c) => c.isRoot && !c.attribute.alias && c.attribute.name === lookup)
		: undefined;
	if (root || key.includes(".")) return root;
	const name = key.toLowerCase();
	return columns.find(
		(c) =>
			!c.isRoot &&
			!c.attribute.alias &&
			(c.container as LinkEntityNode).linkType === "matchfirstrowusingcrossapply" &&
			c.attribute.name === name,
	);
}

export interface RemoveColumnsPlan {
	/** Attribute, order and link-entity node ids to remove */
	removeNodeIds: string[];
	/** Link-entities removed because nothing was left in them */
	removedLinks: LinkEntityNode[];
}

/**
 * The joined link-entity a root `entityname` points at: the one with that alias, else none when it's the
 * root entity's name, else the one with that name. Verified live (2026-11-26, #401): in a self-join, a
 * root order with entityname = the root's name sorts the root entity, not the same-named link-entity.
 */
function linkByEntityname(fetch: FetchNode, links: LinkEntityNode[], entityname: string) {
	const aliased = links.find((l) => l.alias === entityname);
	if (aliased || entityname === fetch.entity.name) return aliased;
	return links.find((l) => l.name === entityname);
}

/**
 * Orders and conditions elsewhere in the query that point at each joined link-entity:
 * - root orders through `entityname`;
 * - conditions anywhere (the root's filters, joined link-entities' filters, link-entities inside
 *   filters, nested filters too) that name it in `entityname`, or compare to it through
 *   `valueof="alias.attribute"` (a valueof qualifier names an alias only).
 */
function references(fetch: FetchNode, links: LinkEntityNode[]) {
	const orders = new Map<string, OrderNode[]>();
	const conditions = new Set<string>();
	for (const order of fetch.entity.orders) {
		const link = order.entityname
			? linkByEntityname(fetch, links, order.entityname)
			: undefined;
		if (link) orders.set(link.id, [...(orders.get(link.id) ?? []), order]);
	}
	const visitFilter = (filter: FilterNode) => {
		for (const condition of filter.conditions) {
			const byName = condition.entityname
				? linkByEntityname(fetch, links, condition.entityname)
				: undefined;
			if (byName) conditions.add(byName.id);
			const dot = condition.valueof?.indexOf(".") ?? -1;
			const qualifier = dot > 0 ? condition.valueof!.slice(0, dot) : undefined;
			const compared = qualifier ? links.find((l) => l.alias === qualifier) : undefined;
			if (compared) conditions.add(compared.id);
		}
		filter.subfilters.forEach(visitFilter);
		filter.links.forEach(visitContainer);
	};
	const visitContainer = (container: EntityNode | LinkEntityNode) => {
		container.filters.forEach(visitFilter);
		container.links.forEach(visitContainer);
	};
	visitContainer(fetch.entity);
	return { orders, conditions };
}

/**
 * What removing columns changes in the tree:
 * - the attributes behind the columns;
 * - orders that sort by a removed attribute's alias (Dataverse rejects an order alias that no
 *   longer exists);
 * - outer link-entities left with nothing in them (no attributes, filters, orders or link-entities):
 *   they would still join, and a one-to-many join keeps multiplying the rows. Inner joins are kept,
 *   because they still filter the rows.
 *
 * Orders and conditions can point at a link-entity from elsewhere: `entityname`, or a condition's
 * `valueof="alias.attribute"` (#357). Owner decisions, recorded on #357:
 * - a condition that names or compares to the link-entity keeps it, because it still filters the rows;
 * - an order that sorts by a removed column of the link-entity (root, or inside the link-entity)
 *   goes with the link-entity, so no order is left pointing at a link-entity that's gone
 *   (Dataverse rejects that: 0x80041108); an order on another attribute keeps the link-entity.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/order
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/condition
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows#cross-table-column-comparisons
 */
export function planRemoveColumns(fetch: FetchNode, keys: readonly string[]): RemoveColumnsPlan {
	const removed = new Set<string>();
	for (const key of keys) {
		const column = resolveColumn(fetch, key);
		if (!column) continue;
		removed.add(column.attribute.id);
		if (column.attribute.alias) {
			findOrdersUsingAlias(fetch, column.container, column.attribute.alias).forEach((o) =>
				removed.add(o.id),
			);
		}
	}
	const refs = references(fetch, collectQueryColumnsLinks(fetch));
	/**
	 * Orders on the link-entity, root ones included, that sort by a removed column's attribute. They only
	 * matter when every attribute of the link-entity is going (otherwise it isn't empty).
	 */
	const columnOrders = (link: LinkEntityNode) => {
		const removedNames = new Set(
			link.attributes.filter((a) => removed.has(a.id)).map((a) => a.name),
		);
		return [...link.orders, ...(refs.orders.get(link.id) ?? [])].filter(
			(o) => !removed.has(o.id) && !o.alias && !!o.attribute && removedNames.has(o.attribute),
		);
	};
	const isEmptyLink = (link: LinkEntityNode) => {
		const going = new Set(columnOrders(link).map((o) => o.id));
		return (
			link.attributes.every((a) => removed.has(a.id)) &&
			!link.allAttributes?.enabled &&
			link.filters.length === 0 &&
			!refs.conditions.has(link.id) &&
			[...link.orders, ...(refs.orders.get(link.id) ?? [])].every(
				(o) => removed.has(o.id) || going.has(o.id),
			) &&
			link.links.every((l) => removed.has(l.id))
		);
	};
	// Innermost first, so a parent left empty by removing its child is removed too
	const removedLinks: LinkEntityNode[] = [];
	const visit = (container: EntityNode | LinkEntityNode) =>
		container.links.forEach((link) => {
			visit(link);
			if (
				link.linkType === "outer" &&
				link.attributes.some((a) => removed.has(a.id)) &&
				isEmptyLink(link)
			) {
				columnOrders(link).forEach((o) => removed.add(o.id));
				removed.add(link.id);
				removedLinks.push(link);
			}
		});
	visit(fetch.entity);
	// Don't list children of removed link-entities separately
	const insideRemovedLink = new Set(
		removedLinks.flatMap((l) => [
			...l.attributes.map((a) => a.id),
			...l.orders.map((o) => o.id),
		]),
	);
	return { removeNodeIds: [...removed].filter((id) => !insideRemovedLink.has(id)), removedLinks };
}

export type SortSpec = Pick<OrderNode, "attribute" | "alias" | "entityname">;

/** Where a header-click sort goes: the root entity's orders, or inside a link-entity. */
export interface SortPlan {
	spec: SortSpec;
	/** The link-entity to hold the order when the root can't target it (no alias, name not unique) */
	insideLink?: LinkEntityNode;
}

/**
 * The order to write when a column header is clicked, or the reason it can't be sorted.
 * - Aggregate queries sort by alias; other queries by attribute name (verified: an attribute alias
 *   fails outside aggregate queries, 0x80041103).
 * - A link-entity column is sorted from the root entity with entityname = the link-entity's alias, or
 *   its name when that's unique; otherwise the order goes inside the link-entity.
 * - Attribute types Dataverse can't sort by aren't sortable (attributeCapabilities, verified matrix).
 */
export function planSortByColumn(
	fetch: FetchNode,
	key: string,
	typeInfo?: AttributeTypeInfo,
): SortPlan | { reason: string } {
	const column = resolveColumn(fetch, key);
	if (!column) return { reason: "This column isn't an attribute of the query." };
	const typeIssue = getOrderMetadataIssue(typeInfo);
	if (typeIssue) return { reason: typeIssue };
	const aggregate = fetch.options.aggregate === true;
	if (aggregate && !column.attribute.alias)
		return { reason: "Aggregate queries sort by alias; give the attribute an alias first." };
	const spec: SortSpec = aggregate
		? { alias: column.attribute.alias }
		: { attribute: column.attribute.name };
	if (column.isRoot) return { spec };
	const link = column.container as LinkEntityNode;
	const sameName = collectQueryColumnsLinks(fetch).filter((l) => l.name === link.name).length;
	const target = link.alias ?? (sameName === 1 ? link.name : undefined);
	return target ? { spec: { ...spec, entityname: target } } : { spec, insideLink: link };
}

function collectQueryColumnsLinks(fetch: FetchNode): LinkEntityNode[] {
	const links: LinkEntityNode[] = [];
	const visit = (container: EntityNode | LinkEntityNode) =>
		container.links.forEach((link) => {
			links.push(link);
			visit(link);
		});
	visit(fetch.entity);
	return links;
}

/** Result columns the query's orders sort by, with their direction (for the grid's sort arrows). */
export function sortedColumns(fetch: FetchNode): Map<string, "ascending" | "descending"> {
	const columns = collectQueryColumns(fetch);
	const links = collectQueryColumnsLinks(fetch);
	const result = new Map<string, "ascending" | "descending">();
	const mark = (order: OrderNode, container: EntityNode | LinkEntityNode) => {
		const direction = order.descending ? "descending" : "ascending";
		for (const column of columns) {
			if (column.container !== container) continue;
			const matches = order.alias
				? column.attribute.alias === order.alias
				: column.attribute.name === order.attribute;
			if (matches && !result.has(column.key)) result.set(column.key, direction);
		}
	};
	for (const order of fetch.entity.orders) {
		const target = order.entityname
			? (links.find((l) => l.alias === order.entityname) ??
				links.find((l) => l.name === order.entityname))
			: fetch.entity;
		if (target) mark(order, target);
	}
	for (const link of links) for (const order of link.orders) mark(order, link);
	return result;
}
