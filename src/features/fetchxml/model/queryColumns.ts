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

import type { AttributeNode, EntityNode, FetchNode, LinkEntityNode, OrderNode } from "./nodes";
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
				columns.push({ key: attribute.alias || `${linkAlias}.${attribute.name}`, attribute, container: link, isRoot: false, linkAlias });
			}
			visit(link);
		});
	visit(fetch.entity);
	return columns;
}

/** The query column behind a result/layout column key (also accepts the _name_value lookup form). */
export function resolveColumn(fetch: FetchNode, key: string): QueryColumn | undefined {
	const columns = collectQueryColumns(fetch);
	const direct = columns.find((c) => c.key === key);
	if (direct) return direct;
	const lookup = /^_(.+)_value$/.exec(key)?.[1];
	return lookup ? columns.find((c) => c.isRoot && !c.attribute.alias && c.attribute.name === lookup) : undefined;
}

export interface RemoveColumnsPlan {
	/** Attribute, order and link-entity node ids to remove */
	removeNodeIds: string[];
	/** Link-entities removed because nothing was left in them */
	removedLinks: LinkEntityNode[];
}

const isEmptyLink = (link: LinkEntityNode, removed: Set<string>) =>
	link.attributes.every((a) => removed.has(a.id)) &&
	!link.allAttributes?.enabled &&
	link.filters.length === 0 &&
	link.orders.every((o) => removed.has(o.id)) &&
	link.links.every((l) => removed.has(l.id));

/**
 * What removing columns changes in the tree:
 * - the attributes behind the columns;
 * - orders that sort by a removed attribute's alias (Dataverse rejects an order alias that no
 *   longer exists);
 * - outer link-entities left with nothing in them (no attributes, filters, orders or link-entities):
 *   they would still join, and a one-to-many join keeps multiplying the rows. Inner joins are kept,
 *   because they still filter the rows.
 */
export function planRemoveColumns(fetch: FetchNode, keys: readonly string[]): RemoveColumnsPlan {
	const removed = new Set<string>();
	for (const key of keys) {
		const column = resolveColumn(fetch, key);
		if (!column) continue;
		removed.add(column.attribute.id);
		if (column.attribute.alias) {
			findOrdersUsingAlias(fetch, column.container, column.attribute.alias).forEach((o) => removed.add(o.id));
		}
	}
	// Innermost first, so a parent left empty by removing its child is removed too
	const removedLinks: LinkEntityNode[] = [];
	const visit = (container: EntityNode | LinkEntityNode) =>
		container.links.forEach((link) => {
			visit(link);
			if (link.linkType === "outer" && isEmptyLink(link, removed) && link.attributes.some((a) => removed.has(a.id))) {
				removed.add(link.id);
				removedLinks.push(link);
			}
		});
	visit(fetch.entity);
	// Don't list children of removed link-entities separately
	const insideRemovedLink = new Set(removedLinks.flatMap((l) => [...l.attributes.map((a) => a.id), ...l.orders.map((o) => o.id)]));
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
	if (aggregate && !column.attribute.alias) return { reason: "Aggregate queries sort by alias; give the attribute an alias first." };
	const spec: SortSpec = aggregate ? { alias: column.attribute.alias } : { attribute: column.attribute.name };
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
			const matches = order.alias ? column.attribute.alias === order.alias : column.attribute.name === order.attribute;
			if (matches && !result.has(column.key)) result.set(column.key, direction);
		}
	};
	for (const order of fetch.entity.orders) {
		const target = order.entityname
			? (links.find((l) => l.alias === order.entityname) ?? links.find((l) => l.name === order.entityname))
			: fetch.entity;
		if (target) mark(order, target);
	}
	for (const link of links) for (const order of link.orders) mark(order, link);
	return result;
}
