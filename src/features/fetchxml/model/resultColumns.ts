/**
 * The results grid's columns (#91), worked out from the query rather than from the keys the first row
 * happens to have: Dataverse leaves null values out of each row, so a column that is null in row 1 (or
 * on the whole first page) has no key there. Shared by the grid and the local Excel export.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/all-attributes
 *
 * - Every attribute of the query is a column, keyed as Dataverse returns it (queryColumns.ts), even when
 *   no loaded row has a value. An unaliased root lookup comes back as the Web API lookup property
 *   _name_value: decided from metadata, or else from the rows.
 *   https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query/select-columns#lookup-property-data
 * - matchfirstrowusingcrossapply returns its attributes by SchemaName, without the link-entity alias;
 *   exists and in link-entities return no columns (attributes in them fail, 0x80041121: fetchxmlRules.ts).
 *   https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables#use-advanced-link-types
 * - Keys no attribute accounts for are columns only where the query asks for all attributes: an
 *   <all-attributes /> element, or a root entity without attribute elements ("If you do not specify
 *   columns, or you use the all-attributes element, data for all columns is returned"). They're
 *   collected from every loaded row, so later pages can add columns but never move or drop them.
 *   https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/select-columns
 */

import type { EntityNode, FetchNode, LinkEntityNode } from "./nodes";
import type { LayoutXmlConfig } from "./layoutxml";
import { collectQueryColumns, joinedLinkAliases, type QueryColumn } from "./queryColumns";
import { NON_RETURNING_LINK_TYPES } from "./fetchxmlRules";

/** What the columns need from attribute metadata */
export interface ResultAttributeInfo {
	AttributeType?: string;
	SchemaName?: string;
}

/** Attribute metadata by entity and attribute logical name (undefined while it's loading) */
export type ResultAttributeLookup = (entityName: string, attributeName: string) => ResultAttributeInfo | undefined;

export interface ResultColumn {
	/** Property name in the rows */
	key: string;
	/** The query attribute behind the column (none for all-attributes columns) */
	column?: QueryColumn;
}

/** Attribute types the Web API returns as a _name_value lookup property */
const LOOKUP_TYPES = new Set(["Lookup", "Customer", "Owner"]);

const CROSS_APPLY = "matchfirstrowusingcrossapply";

/** The rows' data keys (annotations left out), in the order they're first seen */
export function rowKeys(rows: readonly Record<string, unknown>[]): string[] {
	const keys = new Set<string>();
	for (const row of rows) for (const key of Object.keys(row)) if (!key.includes("@")) keys.add(key);
	return [...keys];
}

/** Joined link-entities, and whether each returns columns (not in or under an exists / in link-entity) */
function joinedLinks(fetch: FetchNode): { link: LinkEntityNode; returnsColumns: boolean }[] {
	const result: { link: LinkEntityNode; returnsColumns: boolean }[] = [];
	const visit = (container: EntityNode | LinkEntityNode, returnsColumns: boolean) =>
		container.links.forEach((link) => {
			const returns = returnsColumns && !NON_RETURNING_LINK_TYPES.has(link.linkType);
			result.push({ link, returnsColumns: returns });
			visit(link, returns);
		});
	visit(fetch.entity, true);
	return result;
}

/**
 * The columns of `rows`, in query order, then all-attributes columns in the order first seen.
 * @param executed - the query that produced the rows, when the tree has changed since: attributes
 *   added after the run aren't shown as empty columns until it runs again.
 */
export function resultColumns(
	fetch: FetchNode | null | undefined,
	rows: readonly Record<string, unknown>[],
	getAttribute?: ResultAttributeLookup,
	executed?: FetchNode | null,
): ResultColumn[] {
	const seen = rowKeys(rows);
	if (!fetch) return seen.map((key) => ({ key }));
	const seenSet = new Set(seen);
	const links = joinedLinks(fetch);
	const noColumns = new Set(links.filter((l) => !l.returnsColumns).map((l) => l.link.id));

	const result: ResultColumn[] = [];
	const claimed = new Set<string>();
	const add = (column: ResultColumn) => {
		if (claimed.has(column.key)) return;
		claimed.add(column.key);
		result.push(column);
	};

	const keyOf = (column: QueryColumn): string => {
		const { attribute, container } = column;
		if (attribute.alias) return column.key;
		const info = getAttribute?.(container.name, attribute.name);
		if (column.isRoot) {
			const lookupKey = `_${attribute.name}_value`;
			if (info?.AttributeType) return LOOKUP_TYPES.has(info.AttributeType) ? lookupKey : attribute.name;
			return seenSet.has(lookupKey) ? lookupKey : attribute.name;
		}
		if ((container as LinkEntityNode).linkType === CROSS_APPLY) {
			if (info?.SchemaName) return info.SchemaName;
			// The logical name is the schema name in lowercase:
			// https://learn.microsoft.com/dotnet/api/microsoft.xrm.sdk.metadata.attributemetadata.logicalname
			const found = seen.find((k) => !k.includes(".") && !claimed.has(k) && k.toLowerCase() === attribute.name);
			if (found) return found;
		}
		return column.key;
	};

	// Root columns come first, so a cross-apply key never takes a root attribute's key
	const ran = executed && executed !== fetch ? new Set(resultColumns(executed, rows, getAttribute).map((c) => c.key)) : undefined;
	for (const column of collectQueryColumns(fetch)) {
		if (!column.isRoot && noColumns.has(column.container.id)) continue;
		const key = keyOf(column);
		if (ran && !ran.has(key) && !seenSet.has(key)) continue;
		add({ key, column });
	}

	// All-attributes columns: whatever the rows have for those entities
	const rootAll = !!fetch.entity.allAttributes?.enabled || fetch.entity.attributes.length === 0;
	const aliases = joinedLinkAliases(fetch);
	const allLinks = links.filter((l) => l.returnsColumns && l.link.allAttributes?.enabled).map((l) => l.link);
	const linkPrefixes = allLinks.filter((l) => l.linkType !== CROSS_APPLY).map((l) => `${aliases.get(l.id)}.`);
	// Cross-apply keys are schema names; root and link-entity keys are logical names (all lowercase)
	const crossApplyAll = allLinks.some((l) => l.linkType === CROSS_APPLY);
	for (const key of seen) {
		if (claimed.has(key)) continue;
		const show = key.includes(".")
			? linkPrefixes.some((prefix) => key.startsWith(prefix))
			: rootAll || (crossApplyAll && key !== key.toLowerCase());
		if (show) add({ key });
	}
	return result;
}

/**
 * Columns in the layout's order (#91); columns the layout doesn't have go last, in their own order.
 * Layout cells use the query's column keys, so a column also matches by its attribute's key
 * (_name_value and cross-apply schema names), and an all-attributes _name_value by the name.
 */
export function orderByLayout(columns: ResultColumn[], config: LayoutXmlConfig | null | undefined): ResultColumn[] {
	if (!config || config.columns.length === 0) return columns;
	const index = new Map<string, number>();
	config.columns.forEach((c, i) => {
		if (!index.has(c.name)) index.set(c.name, i);
	});
	const position = (c: ResultColumn): number | undefined => {
		const direct = index.get(c.key) ?? (c.column ? index.get(c.column.key) : undefined);
		if (direct !== undefined) return direct;
		const lookup = /^_(.+)_value$/.exec(c.key)?.[1];
		return lookup ? index.get(lookup) : undefined;
	};
	const inLayout = columns.filter((c) => position(c) !== undefined);
	const rest = columns.filter((c) => position(c) === undefined);
	return [...inLayout.sort((a, b) => position(a)! - position(b)!), ...rest];
}
