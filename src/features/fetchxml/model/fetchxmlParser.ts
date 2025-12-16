/**
 * FetchXML Parser - Convert FetchXML string to FetchNode tree
 * Strict XML syntax validation, lenient schema validation (warnings)
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
	LinkType,
	OperatorType,
	NodeId,
} from "./nodes";

// ID generator for parsed nodes
let parseIdCounter = 0;
const generateParseId = (): NodeId => `parsed_${++parseIdCounter}`;

/**
 * Reset the ID counter (useful for testing)
 */
export function resetParserIdCounter(): void {
	parseIdCounter = 0;
}

/**
 * Parsing result with optional warnings
 */
export interface ParseResult {
	success: boolean;
	fetchNode?: FetchNode;
	errors: ParseError[];
	warnings: ParseWarning[];
}

export interface ParseError {
	message: string;
	line?: number;
	column?: number;
}

export interface ParseWarning {
	message: string;
	element?: string;
	attribute?: string;
}

/**
 * Main entry point: Parse FetchXML string to FetchNode
 */
export function parseFetchXml(xmlString: string): ParseResult {
	const errors: ParseError[] = [];
	const warnings: ParseWarning[] = [];

	// Validate input
	if (!xmlString || typeof xmlString !== "string") {
		return {
			success: false,
			errors: [{ message: "FetchXML string is required" }],
			warnings: [],
		};
	}

	const trimmedXml = xmlString.trim();
	if (!trimmedXml) {
		return {
			success: false,
			errors: [{ message: "FetchXML string is empty" }],
			warnings: [],
		};
	}

	// Parse XML using DOMParser
	let doc: Document;
	try {
		const parser = new DOMParser();
		doc = parser.parseFromString(trimmedXml, "application/xml");

		// Check for parsing errors
		const parserError = doc.querySelector("parsererror");
		if (parserError) {
			// Extract error message from parsererror element
			const errorText = parserError.textContent || "Invalid XML syntax";
			return {
				success: false,
				errors: [{ message: `XML Parsing Error: ${errorText}` }],
				warnings: [],
			};
		}
	} catch (e) {
		return {
			success: false,
			errors: [{ message: `XML Parsing Error: ${e instanceof Error ? e.message : String(e)}` }],
			warnings: [],
		};
	}

	// Validate root element is <fetch>
	const fetchElement = doc.documentElement;
	if (!fetchElement || fetchElement.nodeName.toLowerCase() !== "fetch") {
		return {
			success: false,
			errors: [
				{ message: `Root element must be <fetch>, found <${fetchElement?.nodeName || "none"}>` },
			],
			warnings: [],
		};
	}

	// Find the entity element (required)
	const entityElement = fetchElement.querySelector(":scope > entity");
	if (!entityElement) {
		return {
			success: false,
			errors: [{ message: "Missing required <entity> element inside <fetch>" }],
			warnings: [],
		};
	}

	// Validate entity has a name attribute
	const entityName = entityElement.getAttribute("name");
	if (!entityName) {
		return {
			success: false,
			errors: [{ message: "Entity element must have a 'name' attribute" }],
			warnings: [],
		};
	}

	try {
		// Parse the fetch options
		const options = parseFetchOptions(fetchElement, warnings);

		// Parse the entity tree
		const entity = parseEntity(entityElement, warnings);

		const fetchNode: FetchNode = {
			id: generateParseId(),
			type: "fetch",
			entity,
			options,
		};

		return {
			success: true,
			fetchNode,
			errors,
			warnings,
		};
	} catch (e) {
		return {
			success: false,
			errors: [{ message: `Parsing Error: ${e instanceof Error ? e.message : String(e)}` }],
			warnings,
		};
	}
}

/**
 * Parse fetch element attributes into options
 */
function parseFetchOptions(fetchElement: Element, warnings: ParseWarning[]): FetchNode["options"] {
	const options: FetchNode["options"] = {};

	// Boolean options
	if (fetchElement.getAttribute("aggregate") === "true") {
		options.aggregate = true;
	}
	if (fetchElement.getAttribute("distinct") === "true") {
		options.distinct = true;
	}
	if (fetchElement.getAttribute("returntotalrecordcount") === "true") {
		options.returnTotalRecordCount = true;
	}
	if (fetchElement.getAttribute("no-lock") === "true") {
		options.noLock = true;
	}
	if (fetchElement.getAttribute("latematerialize") === "true") {
		options.latematerialize = true;
	}

	// Numeric options
	const top = fetchElement.getAttribute("top");
	if (top !== null) {
		const topNum = parseInt(top, 10);
		if (!isNaN(topNum)) {
			options.top = topNum;
		} else {
			warnings.push({ message: `Invalid 'top' value: ${top}`, element: "fetch", attribute: "top" });
		}
	}

	const count = fetchElement.getAttribute("count");
	if (count !== null) {
		const countNum = parseInt(count, 10);
		if (!isNaN(countNum)) {
			options.count = countNum;
		} else {
			warnings.push({
				message: `Invalid 'count' value: ${count}`,
				element: "fetch",
				attribute: "count",
			});
		}
	}

	const page = fetchElement.getAttribute("page");
	if (page !== null) {
		const pageNum = parseInt(page, 10);
		if (!isNaN(pageNum)) {
			options.page = pageNum;
		} else {
			warnings.push({
				message: `Invalid 'page' value: ${page}`,
				element: "fetch",
				attribute: "page",
			});
		}
	}

	const utcOffset = fetchElement.getAttribute("utc-offset");
	if (utcOffset !== null) {
		const utcNum = parseInt(utcOffset, 10);
		if (!isNaN(utcNum)) {
			options.utcOffset = utcNum;
		} else {
			warnings.push({
				message: `Invalid 'utc-offset' value: ${utcOffset}`,
				element: "fetch",
				attribute: "utc-offset",
			});
		}
	}

	// String options
	const pagingCookie = fetchElement.getAttribute("paging-cookie");
	if (pagingCookie) {
		options.pagingCookie = pagingCookie;
	}

	// Check for unknown attributes (schema warnings)
	const knownFetchAttrs = new Set([
		"aggregate",
		"distinct",
		"top",
		"count",
		"page",
		"paging-cookie",
		"returntotalrecordcount",
		"no-lock",
		"utc-offset",
		"latematerialize",
		"version",
		"mapping",
		"output-format",
		"min-active-row-version",
		"datasource",
		"options",
	]);

	for (const attr of fetchElement.attributes) {
		if (!knownFetchAttrs.has(attr.name.toLowerCase())) {
			warnings.push({
				message: `Unknown attribute '${attr.name}' on <fetch> element`,
				element: "fetch",
				attribute: attr.name,
			});
		}
	}

	return options;
}

/**
 * Parse entity element
 */
function parseEntity(entityElement: Element, warnings: ParseWarning[]): EntityNode {
	const name = entityElement.getAttribute("name") || "";

	const entity: EntityNode = {
		id: generateParseId(),
		type: "entity",
		name,
		attributes: [],
		orders: [],
		filters: [],
		links: [],
	};

	// Parse optional entity attributes
	if (entityElement.getAttribute("enableprefiltering") === "true") {
		entity.enablePrefiltering = true;
	}
	const prefilterParam = entityElement.getAttribute("prefilterparametername");
	if (prefilterParam) {
		entity.prefilterParameterName = prefilterParam;
	}

	// Parse child elements
	for (const child of entityElement.children) {
		const tagName = child.nodeName.toLowerCase();

		switch (tagName) {
			case "all-attributes":
				entity.allAttributes = parseAllAttributes(child, warnings);
				break;
			case "attribute":
				entity.attributes.push(parseAttribute(child, warnings));
				break;
			case "order":
				entity.orders.push(parseOrder(child, warnings));
				break;
			case "filter":
				entity.filters.push(parseFilter(child, warnings));
				break;
			case "link-entity":
				entity.links.push(parseLinkEntity(child, warnings));
				break;
			default:
				warnings.push({
					message: `Unknown element <${tagName}> inside <entity>`,
					element: tagName,
				});
		}
	}

	return entity;
}

/**
 * Parse all-attributes element
 * Note: Parameters reserved for future use (potential warnings for deprecated usage)
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function parseAllAttributes(_element: Element, _warnings: ParseWarning[]): AllAttributesNode {
	return {
		id: generateParseId(),
		type: "all-attributes",
		enabled: true,
	};
}

/**
 * Parse attribute element
 */
function parseAttribute(attrElement: Element, warnings: ParseWarning[]): AttributeNode {
	const name = attrElement.getAttribute("name") || "";

	if (!name) {
		warnings.push({
			message: "Attribute element missing 'name'",
			element: "attribute",
		});
	}

	const attr: AttributeNode = {
		id: generateParseId(),
		type: "attribute",
		name,
	};

	// Optional attributes
	const alias = attrElement.getAttribute("alias");
	if (alias) attr.alias = alias;

	if (attrElement.getAttribute("groupby") === "true") {
		attr.groupby = true;
	}

	const aggregate = attrElement.getAttribute("aggregate");
	if (aggregate) {
		const validAggregates = ["sum", "count", "countcolumn", "min", "max", "avg", "rowaggregate"];
		if (validAggregates.includes(aggregate)) {
			attr.aggregate = aggregate as AttributeNode["aggregate"];
		} else {
			warnings.push({
				message: `Invalid aggregate value '${aggregate}'`,
				element: "attribute",
				attribute: "aggregate",
			});
		}
	}

	const dategrouping = attrElement.getAttribute("dategrouping");
	if (dategrouping) {
		const validDateGroupings = [
			"day",
			"week",
			"month",
			"quarter",
			"year",
			"fiscal-period",
			"fiscal-year",
		];
		if (validDateGroupings.includes(dategrouping)) {
			attr.dategrouping = dategrouping as AttributeNode["dategrouping"];
		} else {
			warnings.push({
				message: `Invalid dategrouping value '${dategrouping}'`,
				element: "attribute",
				attribute: "dategrouping",
			});
		}
	}

	if (attrElement.hasAttribute("usertimezone")) {
		attr.usertimezone = attrElement.getAttribute("usertimezone") === "true";
	}

	return attr;
}

/**
 * Parse order element
 */
function parseOrder(orderElement: Element, warnings: ParseWarning[]): OrderNode {
	const attribute = orderElement.getAttribute("attribute") || "";

	if (!attribute) {
		warnings.push({
			message: "Order element missing 'attribute'",
			element: "order",
		});
	}

	const order: OrderNode = {
		id: generateParseId(),
		type: "order",
		attribute,
	};

	if (orderElement.getAttribute("descending") === "true") {
		order.descending = true;
	}

	const entityname = orderElement.getAttribute("entityname");
	if (entityname) {
		order.entityname = entityname;
	}

	return order;
}

/**
 * Parse filter element (recursive for subfilters)
 */
function parseFilter(filterElement: Element, warnings: ParseWarning[]): FilterNode {
	const type = filterElement.getAttribute("type") || "and";

	const filter: FilterNode = {
		id: generateParseId(),
		type: "filter",
		conjunction: type === "or" ? "or" : "and",
		conditions: [],
		subfilters: [],
		links: [],
	};

	const hint = filterElement.getAttribute("hint");
	if (hint === "union") {
		filter.hint = "union";
	}

	// Parse child elements
	for (const child of filterElement.children) {
		const tagName = child.nodeName.toLowerCase();

		switch (tagName) {
			case "condition":
				filter.conditions.push(parseCondition(child, warnings));
				break;
			case "filter":
				filter.subfilters.push(parseFilter(child, warnings));
				break;
			case "link-entity":
				// Parse link-entity inside filter (for any/not any/all/not all)
				filter.links.push(parseLinkEntity(child, warnings));
				break;
			default:
				warnings.push({
					message: `Unknown element <${tagName}> inside <filter>`,
					element: tagName,
				});
		}
	}

	return filter;
}

/**
 * Parse condition element
 */
function parseCondition(condElement: Element, warnings: ParseWarning[]): ConditionNode {
	const attribute = condElement.getAttribute("attribute") || "";
	const operator = condElement.getAttribute("operator") || "eq";

	if (!attribute) {
		warnings.push({
			message: "Condition element missing 'attribute'",
			element: "condition",
		});
	}

	const condition: ConditionNode = {
		id: generateParseId(),
		type: "condition",
		attribute,
		operator: operator as OperatorType,
	};

	// Optional attributes
	const entityname = condElement.getAttribute("entityname");
	if (entityname) condition.entityname = entityname;

	const valueof = condElement.getAttribute("valueof");
	if (valueof) condition.valueof = valueof;

	const aggregate = condElement.getAttribute("aggregate");
	if (aggregate) {
		const validAggregates = ["sum", "count", "countcolumn", "min", "max", "avg"];
		if (validAggregates.includes(aggregate)) {
			condition.aggregate = aggregate as ConditionNode["aggregate"];
		}
	}

	// Parse value - can be from attribute or child <value> elements
	const valueAttr = condElement.getAttribute("value");
	const valueElements = condElement.querySelectorAll(":scope > value");

	if (valueElements.length > 0) {
		// Multiple value elements (for in, not-in, between, etc.)
		const values: (string | number)[] = [];
		for (const valEl of valueElements) {
			const valText = valEl.textContent?.trim() || "";
			// Try to parse as number if it looks like one
			const numVal = parseFloat(valText);
			if (!isNaN(numVal) && valText === numVal.toString()) {
				values.push(numVal);
			} else {
				values.push(valText);
			}
		}
		condition.value = values;
	} else if (valueAttr !== null) {
		// Single value attribute
		const numVal = parseFloat(valueAttr);
		if (!isNaN(numVal) && valueAttr === numVal.toString()) {
			condition.value = numVal;
		} else if (valueAttr === "true") {
			condition.value = true;
		} else if (valueAttr === "false") {
			condition.value = false;
		} else {
			condition.value = valueAttr;
		}
	}

	return condition;
}

/**
 * Parse link-entity element (recursive)
 */
function parseLinkEntity(linkElement: Element, warnings: ParseWarning[]): LinkEntityNode {
	const name = linkElement.getAttribute("name") || "";
	const from = linkElement.getAttribute("from") || "";
	const to = linkElement.getAttribute("to") || "";

	if (!name) {
		warnings.push({
			message: "Link-entity element missing 'name'",
			element: "link-entity",
		});
	}
	if (!from) {
		warnings.push({
			message: "Link-entity element missing 'from'",
			element: "link-entity",
		});
	}
	if (!to) {
		warnings.push({
			message: "Link-entity element missing 'to'",
			element: "link-entity",
		});
	}

	const linkTypeAttr = linkElement.getAttribute("link-type") || "inner";
	const validLinkTypes: LinkType[] = [
		"inner",
		"outer",
		"any",
		"not any",
		"all",
		"not all",
		"exists",
		"in",
		"matchfirstrowusingcrossapply",
	];

	let linkType: LinkType = "inner";
	if (validLinkTypes.includes(linkTypeAttr as LinkType)) {
		linkType = linkTypeAttr as LinkType;
	} else {
		warnings.push({
			message: `Invalid link-type '${linkTypeAttr}', defaulting to 'inner'`,
			element: "link-entity",
			attribute: "link-type",
		});
	}

	const link: LinkEntityNode = {
		id: generateParseId(),
		type: "link-entity",
		name,
		from,
		to,
		linkType,
		attributes: [],
		orders: [],
		filters: [],
		links: [],
	};

	// Optional attributes
	const alias = linkElement.getAttribute("alias");
	if (alias) link.alias = alias;

	if (linkElement.getAttribute("intersect") === "true") {
		link.intersect = true;
	}

	if (linkElement.hasAttribute("visible")) {
		link.visible = linkElement.getAttribute("visible") === "true";
	}

	// Parse child elements
	for (const child of linkElement.children) {
		const tagName = child.nodeName.toLowerCase();

		switch (tagName) {
			case "all-attributes":
				link.allAttributes = parseAllAttributes(child, warnings);
				break;
			case "attribute":
				link.attributes.push(parseAttribute(child, warnings));
				break;
			case "order":
				link.orders.push(parseOrder(child, warnings));
				break;
			case "filter":
				link.filters.push(parseFilter(child, warnings));
				break;
			case "link-entity":
				link.links.push(parseLinkEntity(child, warnings));
				break;
			default:
				warnings.push({
					message: `Unknown element <${tagName}> inside <link-entity>`,
					element: tagName,
				});
		}
	}

	return link;
}

/**
 * Validate a FetchXML string without fully parsing it
 * Quick syntax check
 */
export function validateFetchXmlSyntax(xmlString: string): { valid: boolean; error?: string } {
	if (!xmlString?.trim()) {
		return { valid: false, error: "FetchXML string is empty" };
	}

	try {
		const parser = new DOMParser();
		const doc = parser.parseFromString(xmlString.trim(), "application/xml");
		const parserError = doc.querySelector("parsererror");
		if (parserError) {
			return { valid: false, error: parserError.textContent || "Invalid XML syntax" };
		}

		const fetchElement = doc.documentElement;
		if (!fetchElement || fetchElement.nodeName.toLowerCase() !== "fetch") {
			return { valid: false, error: "Root element must be <fetch>" };
		}

		const entityElement = fetchElement.querySelector(":scope > entity");
		if (!entityElement) {
			return { valid: false, error: "Missing required <entity> element" };
		}

		if (!entityElement.getAttribute("name")) {
			return { valid: false, error: "Entity element must have a 'name' attribute" };
		}

		return { valid: true };
	} catch (e) {
		return { valid: false, error: e instanceof Error ? e.message : String(e) };
	}
}
