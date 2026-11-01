/**
 * Pure functions to generate FetchXML from tree structure
 * Follows Dataverse FetchXML semantics
 */

import type {
	FetchNode,
	EntityNode,
	AttributeNode,
	AllAttributesNode,
	OrderNode,
	FilterNode,
	ConditionNode,
	LinkEntityNode,
	ExtraAttributes,
} from "./nodes";
import { operatorRequiresValue } from "./operators";
import { AGGREGATE_FUNCTIONS } from "./attributeCapabilities";

/**
 * Convert a value to string for FetchXML, avoiding scientific notation for numbers
 */
function formatValueForFetchXml(val: unknown): string {
	if (typeof val === "number") {
		// String() is exact and round-trips (0.12345678901234 used to be cut to 10 places, #181);
		// only scientific notation (1e-12, 1e21) is written out in full
		const text = String(val);
		return /e/i.test(text) ? val.toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: 20 }) : text;
	}
	return String(val);
}

/** Attributes kept verbatim from parsed FetchXML (#90): unknown, namespaced, legacy. */
function extraAttrs(extra: ExtraAttributes | undefined): string[] {
	return extra ? Object.entries(extra).map(([name, value]) => `${name}="${escapeXmlAttr(value)}"`) : [];
}

/**
 * Generate complete FetchXML from FetchNode.
 *
 * The output reflects exactly what is in the tree — no attributes are added. Dataverse already
 * returns the root table's primary key for non-aggregate, non-distinct queries, and adding it
 * would break aggregate queries and defeat distinct (#54, #59).
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/select-columns
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results#ordering-and-paging
 */
export function generateFetchXml(fetchNode: FetchNode): string {
	const lines: string[] = [];

	// Build fetch element with options
	const fetchAttrs: string[] = [];
	if (fetchNode.options.aggregate) fetchAttrs.push('aggregate="true"');
	if (fetchNode.options.distinct) fetchAttrs.push('distinct="true"');
	if (fetchNode.options.top !== undefined) fetchAttrs.push(`top="${fetchNode.options.top}"`);
	if (fetchNode.options.count !== undefined) fetchAttrs.push(`count="${fetchNode.options.count}"`);
	if (fetchNode.options.page !== undefined) fetchAttrs.push(`page="${fetchNode.options.page}"`);
	if (fetchNode.options.returnTotalRecordCount) fetchAttrs.push('returntotalrecordcount="true"');
	if (fetchNode.options.noLock) fetchAttrs.push('no-lock="true"');
	if (fetchNode.options.utcOffset !== undefined)
		fetchAttrs.push(`utc-offset="${fetchNode.options.utcOffset}"`);
	if (fetchNode.options.pagingCookie)
		fetchAttrs.push(`paging-cookie="${escapeXmlAttr(fetchNode.options.pagingCookie)}"`);
	if (fetchNode.options.aggregateLimit !== undefined)
		fetchAttrs.push(`aggregatelimit="${fetchNode.options.aggregateLimit}"`);
	if (fetchNode.options.dataSource)
		fetchAttrs.push(`datasource="${escapeXmlAttr(fetchNode.options.dataSource)}"`);
	if (fetchNode.options.latematerialize) fetchAttrs.push('latematerialize="true"');
	if (fetchNode.options.useRawOrderBy) fetchAttrs.push('useraworderby="true"');
	if (fetchNode.options.queryHints?.length)
		fetchAttrs.push(`options="${escapeXmlAttr(fetchNode.options.queryHints.join(","))}"`);
	fetchAttrs.push(...extraAttrs(fetchNode.extraAttributes));

	const fetchAttrStr = fetchAttrs.length > 0 ? " " + fetchAttrs.join(" ") : "";
	lines.push(`<fetch${fetchAttrStr}>`);

	// Generate entity
	lines.push(...generateEntity(fetchNode.entity, 1));

	lines.push("</fetch>");

	return lines.join("\n");
}

/**
 * Generate entity element
 */
function generateEntity(entity: EntityNode, indent: number): string[] {
	const lines: string[] = [];
	const spaces = "  ".repeat(indent);

	const entityAttrs: string[] = [`name="${escapeXmlAttr(entity.name)}"`];
	// Documented form for report prefiltering: "You must set the value of the enableprefiltering parameter to 1"
	if (entity.enablePrefiltering) entityAttrs.push('enableprefiltering="1"');
	if (entity.prefilterParameterName)
		entityAttrs.push(`prefilterparametername="${escapeXmlAttr(entity.prefilterParameterName)}"`);
	entityAttrs.push(...extraAttrs(entity.extraAttributes));
	lines.push(`${spaces}<entity ${entityAttrs.join(" ")}>`);

	// All-attributes (single node, not array)
	if (entity.allAttributes?.enabled) {
		lines.push(...generateAllAttributes(entity.allAttributes, indent + 1));
	}

	// Attributes
	entity.attributes.forEach((attr) => {
		lines.push(...generateAttribute(attr, indent + 1));
	});

	// Orders
	entity.orders?.forEach((order) => {
		lines.push(...generateOrder(order, indent + 1));
	});

	// Filters
	entity.filters?.forEach((filter) => {
		lines.push(...generateFilter(filter, indent + 1));
	});

	// Link-entities
	entity.links?.forEach((link) => {
		lines.push(...generateLinkEntity(link, indent + 1));
	});

	lines.push(`${spaces}</entity>`);

	return lines;
}

/**
 * Generate all-attributes element
 */
function generateAllAttributes(_allAttr: AllAttributesNode, indent: number): string[] {
	const spaces = "  ".repeat(indent);
	return [`${spaces}<all-attributes />`];
}

/** Documented values (FetchXML attribute reference); anything else is never written (#88). */
const AGGREGATES: ReadonlySet<string> = new Set(AGGREGATE_FUNCTIONS);
const DATE_GROUPINGS: ReadonlySet<string> = new Set(["day", "week", "month", "quarter", "year", "fiscal-period", "fiscal-year"]);

/**
 * Generate attribute element
 */
function generateAttribute(attr: AttributeNode, indent: number): string[] {
	const spaces = "  ".repeat(indent);
	const attrs: string[] = [`name="${escapeXmlAttr(attr.name)}"`];

	if (attr.alias) attrs.push(`alias="${escapeXmlAttr(attr.alias)}"`);
	if (attr.aggregate && AGGREGATES.has(attr.aggregate)) attrs.push(`aggregate="${attr.aggregate}"`);
	if (attr.groupby) attrs.push('groupby="true"');
	if (attr.dategrouping && DATE_GROUPINGS.has(attr.dategrouping))
		attrs.push(`dategrouping="${attr.dategrouping}"`);
	if (attr.usertimezone !== undefined) attrs.push(`usertimezone="${attr.usertimezone}"`);
	if (attr.distinct) attrs.push('distinct="true"');
	if (attr.rowAggregate) attrs.push(`rowaggregate="${escapeXmlAttr(attr.rowAggregate)}"`);
	attrs.push(...extraAttrs(attr.extraAttributes));

	return [`${spaces}<attribute ${attrs.join(" ")} />`];
}

/**
 * Generate order element
 */
function generateOrder(order: OrderNode, indent: number): string[] {
	const spaces = "  ".repeat(indent);
	// Non-aggregate queries order by attribute; aggregate queries order by alias (#90)
	const attrs: string[] = [];
	if (order.attribute) attrs.push(`attribute="${escapeXmlAttr(order.attribute)}"`);
	if (order.alias) attrs.push(`alias="${escapeXmlAttr(order.alias)}"`);

	if (order.descending) attrs.push('descending="true"');
	if (order.entityname) attrs.push(`entityname="${escapeXmlAttr(order.entityname)}"`);
	attrs.push(...extraAttrs(order.extraAttributes));

	return [`${spaces}<order ${attrs.join(" ")} />`];
}

/**
 * Generate filter element (recursive for subfilters)
 */
function generateFilter(filter: FilterNode, indent: number): string[] {
	const lines: string[] = [];
	const spaces = "  ".repeat(indent);

	const attrs: string[] = [`type="${filter.conjunction}"`];
	if (filter.hint) attrs.push(`hint="${escapeXmlAttr(filter.hint)}"`);
	// Quick find flags use the documented "1" form
	if (filter.isQuickFindFields) attrs.push('isquickfindfields="1"');
	if (filter.overrideQuickFindRecordLimitEnabled) attrs.push('overridequickfindrecordlimitenabled="1"');
	if (filter.overrideQuickFindRecordLimitDisabled) attrs.push('overridequickfindrecordlimitdisabled="1"');
	attrs.push(...extraAttrs(filter.extraAttributes));

	lines.push(`${spaces}<filter ${attrs.join(" ")}>`);

	// Conditions
	filter.conditions?.forEach((condition) => {
		lines.push(...generateCondition(condition, indent + 1));
	});

	// Subfilters (recursive)
	filter.subfilters?.forEach((subfilter) => {
		lines.push(...generateFilter(subfilter, indent + 1));
	});

	// Link-entities inside filter (for any/not any/all/not all)
	filter.links?.forEach((link) => {
		lines.push(...generateLinkEntity(link, indent + 1));
	});

	lines.push(`${spaces}</filter>`);

	return lines;
}

/**
 * Generate condition element
 */
function generateCondition(condition: ConditionNode, indent: number): string[] {
	const lines: string[] = [];
	const spaces = "  ".repeat(indent);

	const attrs: string[] = [
		`attribute="${escapeXmlAttr(condition.attribute)}"`,
		`operator="${escapeXmlAttr(condition.operator)}"`,
	];

	if (condition.entityname) attrs.push(`entityname="${escapeXmlAttr(condition.entityname)}"`);
	if (condition.aggregate && AGGREGATES.has(condition.aggregate))
		attrs.push(`aggregate="${condition.aggregate}"`);
	if (condition.valueof) attrs.push(`valueof="${escapeXmlAttr(condition.valueof)}"`);
	attrs.push(...extraAttrs(condition.extraAttributes));

	// Use the centralized operator logic to determine if value is needed
	const needsValue = operatorRequiresValue(condition.operator);

	if (!needsValue) {
		// Self-closing tag for operators that don't need values (null, eq-userid, etc.)
		return [`${spaces}<condition ${attrs.join(" ")} />`];
	}

	// Operators that need value attribute
	if (condition.value !== undefined && condition.value !== null) {
		// Handle array values (for 'in', 'not-in', 'contain-values', 'between', etc.)
		if (Array.isArray(condition.value)) {
			lines.push(`${spaces}<condition ${attrs.join(" ")}>`);
			condition.value.forEach((val) => {
				if (val !== undefined && val !== null) {
					const text = formatValueForFetchXml(val);
					// The value's own attributes (uiname, uitype) as they were parsed (#181)
					const own = extraAttrs(condition.valueAttributes?.[text]);
					lines.push(`${spaces}  <value${own.length ? ` ${own.join(" ")}` : ""}>${escapeXmlText(text)}</value>`);
				}
			});
			lines.push(`${spaces}</condition>`);
		} else {
			// Single value attribute
			attrs.push(`value="${escapeXmlAttr(formatValueForFetchXml(condition.value))}"`);
			lines.push(`${spaces}<condition ${attrs.join(" ")} />`);
		}
	} else {
		// No value provided but operator needs one - still generate with empty value
		lines.push(`${spaces}<condition ${attrs.join(" ")} />`);
	}

	return lines;
}

/**
 * Generate link-entity element (recursive)
 */
function generateLinkEntity(link: LinkEntityNode, indent: number): string[] {
	const lines: string[] = [];
	const spaces = "  ".repeat(indent);

	// Empty from/to are omitted, not written as from="" (which silently returns no rows); Dataverse
	// infers the join only for system many-to-many relationships.
	const attrs: string[] = [`name="${escapeXmlAttr(link.name)}"`];
	if (link.from) attrs.push(`from="${escapeXmlAttr(link.from)}"`);
	if (link.to) attrs.push(`to="${escapeXmlAttr(link.to)}"`);

	if (link.alias) attrs.push(`alias="${escapeXmlAttr(link.alias)}"`);
	if (link.linkType !== "inner") attrs.push(`link-type="${link.linkType}"`);
	if (link.intersect) attrs.push('intersect="true"');
	if (link.visible !== undefined) attrs.push(`visible="${link.visible}"`);
	attrs.push(...extraAttrs(link.extraAttributes));

	lines.push(`${spaces}<link-entity ${attrs.join(" ")}>`);

	// All-attributes (single node, not array)
	if (link.allAttributes?.enabled) {
		lines.push(...generateAllAttributes(link.allAttributes, indent + 1));
	}

	// Attributes
	link.attributes?.forEach((attr) => {
		lines.push(...generateAttribute(attr, indent + 1));
	});

	// Orders
	link.orders?.forEach((order) => {
		lines.push(...generateOrder(order, indent + 1));
	});

	// Filters
	link.filters?.forEach((filter) => {
		lines.push(...generateFilter(filter, indent + 1));
	});

	// Nested link-entities (recursive)
	link.links?.forEach((nestedLink) => {
		lines.push(...generateLinkEntity(nestedLink, indent + 1));
	});

	lines.push(`${spaces}</link-entity>`);

	return lines;
}

/**
 * Escape XML special characters.
 *
 * Attribute values are always written in double quotes, so `'` stays as is: `&apos;` is valid XML
 * but breaks expressions such as @{body('x')} when the FetchXML is pasted into Power Automate (#63).
 *
 * A missing value is written as empty, never thrown on: the property editors clear required fields
 * while the user is mid-edit (picking a condition's attribute clears its operator, changing its
 * entity clears its attribute), and the FetchXML preview serializes the tree on every render. In
 * 1.4.1 and 1.4.2, `undefined.replace` here crashed the whole tool (1.4.3 hotfix). The incomplete
 * node then shows as attribute="" / operator="" and its editor still asks for the missing value.
 */
function escapeXmlAttr(str: string | null | undefined): string {
	return (
		(str ?? "")
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;")
			// XML attribute-value normalisation turns a literal tab, newline or CR into a space, so
			// they're written as character references to survive (#181)
			.replace(/\t/g, "&#9;")
			.replace(/\n/g, "&#10;")
			.replace(/\r/g, "&#13;")
	);
}

/** Element text (<value>): only &, < and > need escaping. A missing value is written as empty. */
function escapeXmlText(str: string | null | undefined): string {
	return (str ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/**
 * Set the paging attributes on a FetchXML query for the next page (#58, #60):
 * - page: the page to request;
 * - paging-cookie: the decoded cookie from the previous page, XML-escaped here; omitted when
 *   Dataverse returned none (simple paging);
 * - count: the page size; when not given, the query's own count is kept.
 * Existing page / paging-cookie (and count, when given) are replaced, in either quote style.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results
 */
export function addPagingToFetchXml(
	fetchXml: string,
	page: number,
	pagingCookie?: string,
	count?: number,
): string {
	// A quoted value may contain a raw ">" (#200)
	const fetchMatch = /<fetch\b((?:[^>"']|"[^"]*"|'[^']*')*)>/.exec(fetchXml);
	if (!fetchMatch) return fetchXml;

	const names = count !== undefined ? "page|paging-cookie|count" : "page|paging-cookie";
	let fetchAttrs = fetchMatch[1]
		.replace(new RegExp(`\\s+(?:${names})\\s*=\\s*(?:"[^"]*"|'[^']*')`, "g"), "")
		.trim();

	fetchAttrs += ` page="${page}"`;
	if (pagingCookie) fetchAttrs += ` paging-cookie="${escapeXmlAttr(pagingCookie)}"`;
	if (count !== undefined) fetchAttrs += ` count="${count}"`;

	// A function replacement: the cookie holds record values, and "$&" etc. in a string replacement
	// would be expanded
	return fetchXml.replace(fetchMatch[0], () => `<fetch ${fetchAttrs.trim()}>`);
}
