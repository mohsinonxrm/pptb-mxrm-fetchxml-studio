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
} from "./nodes";
import { operatorRequiresValue } from "./operators";

/**
 * Convert a value to string for FetchXML, avoiding scientific notation for numbers
 */
function formatValueForFetchXml(val: unknown): string {
	if (typeof val === "number") {
		// Avoid scientific notation (1e-12) by using toFixed with sufficient precision
		// Then remove trailing zeros
		if (Math.abs(val) < 1e-10 && val !== 0) {
			// Very small numbers - use high precision
			return val.toFixed(20).replace(/\.?0+$/, "");
		} else if (Number.isInteger(val)) {
			// Integer - no decimal point needed
			return val.toString();
		} else {
			// Decimal - preserve precision but remove trailing zeros
			return val.toFixed(10).replace(/\.?0+$/, "");
		}
	}
	return String(val);
}

/**
 * Generate complete FetchXML from FetchNode
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
		fetchAttrs.push(`paging-cookie="${escapeXml(fetchNode.options.pagingCookie)}"`);

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

	lines.push(`${spaces}<entity name="${entity.name}">`);

	// All-attributes (single node, not array)
	if (entity.allAttributes?.enabled) {
		lines.push(...generateAllAttributes(entity.allAttributes, indent + 1));
	}

	// CRITICAL: Always include the primary ID attribute for row selection to work
	// The primary ID is {entityname}id (e.g., accountid, contactid)
	const primaryIdAttrName = `${entity.name}id`;
	const hasPrimaryId =
		entity.attributes.some((attr) => attr.name === primaryIdAttrName) ||
		entity.allAttributes?.enabled;

	if (!hasPrimaryId) {
		// Inject primary ID attribute at the beginning
		const primaryIdSpaces = "  ".repeat(indent + 1);
		lines.push(`${primaryIdSpaces}<attribute name="${primaryIdAttrName}" />`);
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

/**
 * Generate attribute element
 */
function generateAttribute(attr: AttributeNode, indent: number): string[] {
	const spaces = "  ".repeat(indent);
	const attrs: string[] = [`name="${attr.name}"`];

	if (attr.alias) attrs.push(`alias="${attr.alias}"`);
	if (attr.aggregate) attrs.push(`aggregate="${attr.aggregate}"`);
	if (attr.groupby) attrs.push('groupby="true"');
	if (attr.dategrouping) attrs.push(`dategrouping="${attr.dategrouping}"`);
	if (attr.usertimezone !== undefined) attrs.push(`usertimezone="${attr.usertimezone}"`);

	return [`${spaces}<attribute ${attrs.join(" ")} />`];
}

/**
 * Generate order element
 */
function generateOrder(order: OrderNode, indent: number): string[] {
	const spaces = "  ".repeat(indent);
	const attrs: string[] = [`attribute="${order.attribute}"`];

	if (order.descending) attrs.push('descending="true"');
	if (order.entityname) attrs.push(`entityname="${order.entityname}"`);

	return [`${spaces}<order ${attrs.join(" ")} />`];
}

/**
 * Generate filter element (recursive for subfilters)
 */
function generateFilter(filter: FilterNode, indent: number): string[] {
	const lines: string[] = [];
	const spaces = "  ".repeat(indent);

	const attrs: string[] = [`type="${filter.conjunction}"`];
	if (filter.hint) attrs.push(`hint="${filter.hint}"`);

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
		`attribute="${condition.attribute}"`,
		`operator="${condition.operator}"`,
	];

	if (condition.entityname) attrs.push(`entityname="${condition.entityname}"`);
	if (condition.aggregate) attrs.push(`aggregate="${condition.aggregate}"`);
	// Note: ConditionNode doesn't have alias in our type definition
	if (condition.valueof) attrs.push(`valueof="${condition.valueof}"`);

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
					lines.push(`${spaces}  <value>${escapeXml(formatValueForFetchXml(val))}</value>`);
				}
			});
			lines.push(`${spaces}</condition>`);
		} else {
			// Single value attribute
			attrs.push(`value="${escapeXml(formatValueForFetchXml(condition.value))}"`);
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

	const attrs: string[] = [`name="${link.name}"`, `from="${link.from}"`, `to="${link.to}"`];

	if (link.alias) attrs.push(`alias="${link.alias}"`);
	if (link.linkType !== "inner") attrs.push(`link-type="${link.linkType}"`);
	if (link.intersect) attrs.push('intersect="true"');
	if (link.visible !== undefined) attrs.push(`visible="${link.visible}"`);

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
 * Escape XML special characters
 */
function escapeXml(str: string): string {
	return str
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;")
		.replace(/'/g, "&apos;");
}
