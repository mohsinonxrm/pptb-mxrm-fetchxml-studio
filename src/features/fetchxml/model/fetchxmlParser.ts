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
	ExtraAttributes,
} from "./nodes";
import { isValidAlias, sanitizeAlias } from "./aliasUtils";
import { KNOWN_OPERATORS } from "./operators";
import { validateQueryRules } from "./fetchxmlRules";

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

/**
 * How serious a parser finding is (#89). Principle: what the FetchXML reference documents is
 * supported; undocumented behaviour verified to work is supported too; only what is known not to
 * work is blocked. Every classification was validated against a live Dataverse org.
 *  - "error":   known to be rejected by Dataverse — Execute is blocked. Parse to tree can still
 *               build the tree, correcting or dropping the item where possible.
 *  - "warning": can't be determined client-side (e.g. a link-entity without from/to is valid only
 *               for system many-to-many relationships) — the editor offers "Execute anyway".
 *  - "info":    never blocks — accepted or ignored by Dataverse, but worth knowing.
 * Problems that prevent building a tree at all are reported in `errors` instead.
 */
export type ParseWarningSeverity = "error" | "warning" | "info";

/**
 * What happens to the finding:
 *  - "invalid":   the query is invalid as written (Dataverse rejects it)
 *  - "ignored":   Dataverse ignores it; Parse to tree drops it with no change in results
 *  - "dropped":   meaningful, but the tree can't represent it yet — Parse to tree removes it
 *  - "corrected": Parse to tree changes it to a valid form (e.g. sanitized alias)
 *  - "advice":    valid as written and kept as-is; a recommendation only
 */
export type ParseWarningKind = "invalid" | "ignored" | "dropped" | "corrected" | "advice";

export interface ParseWarning {
	message: string;
	element?: string;
	attribute?: string;
	severity: ParseWarningSeverity;
	kind: ParseWarningKind;
	/** Tree node the finding refers to (query-level rules only, #90b) */
	nodeId?: NodeId;
}

/** Internal: "fatal" findings prevent building a tree and are moved to `errors`. */
type ParseIssue = Omit<ParseWarning, "severity"> & { severity: ParseWarningSeverity | "fatal" };

/** Findings Dataverse is known to reject — Execute is blocked until they're fixed. */
export function getExecutionErrors(result: Pick<ParseResult, "warnings">): ParseWarning[] {
	return result.warnings.filter((w) => w.severity === "error");
}

/** Findings that can't be determined client-side — the editor offers "Execute anyway". */
export function getBlockingWarnings(result: Pick<ParseResult, "warnings">): ParseWarning[] {
	return result.warnings.filter((w) => w.severity === "warning");
}

/**
 * Changes Parse to tree would make to the query: anything it drops, and anything it rewrites to fix
 * a problem Dataverse would reject. Ignored items (Dataverse ignores them too) and minor
 * informational corrections are not changes; invalid items are kept as written.
 */
export function getParseToTreeChanges(result: Pick<ParseResult, "warnings">): ParseWarning[] {
	return result.warnings.filter(
		(w) => w.kind === "dropped" || (w.kind === "corrected" && w.severity !== "info"),
	);
}

const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";


/**
 * A whole number as Dataverse reads top, count, page and utc-offset: optional spaces and sign,
 * digits only (#181). parseInt accepted "10abc" as 10. Verified live (2026-10-30): top="10abc"
 * fails ("Input string was not in a correct format"), top=" 10 " and top="+5" run.
 */
function parseXmlInt(text: string): number {
	return /^\s*[+-]?\d+\s*$/.test(text) ? parseInt(text, 10) : NaN;
}

/**
 * Attributes Dataverse ignores that are ubiquitous in FetchXML exported from Advanced Find and
 * saved views. Reporting them on every paste would be noise, so they are accepted silently.
 */
const SILENT_LEGACY_ATTRIBUTES: Record<string, ReadonlySet<string>> = {
	fetch: new Set(["version", "output-format", "mapping", "min-active-row-version"]),
	condition: new Set(["uiname", "uitype", "uihidden"]),
};

/**
 * Per element: attributes with a dedicated field in the tree model. Every attribute in the FetchXML
 * reference is modelled (#90); anything else is kept verbatim in `extraAttributes`.
 * Source: https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/
 */
const MODELLED_ATTRIBUTES: Record<string, ReadonlySet<string>> = {
	fetch: new Set([
		"aggregate",
		"aggregatelimit",
		"count",
		"datasource",
		"distinct",
		"latematerialize",
		"no-lock",
		"options",
		"page",
		"paging-cookie",
		"returntotalrecordcount",
		"top",
		"useraworderby",
		"utc-offset",
	]),
	entity: new Set(["name", "enableprefiltering", "prefilterparametername"]),
	attribute: new Set([
		"name",
		"alias",
		"aggregate",
		"dategrouping",
		"distinct",
		"groupby",
		"rowaggregate",
		"usertimezone",
	]),
	order: new Set(["attribute", "alias", "descending", "entityname"]),
	filter: new Set([
		"type",
		"hint",
		"isquickfindfields",
		"overridequickfindrecordlimitenabled",
		"overridequickfindrecordlimitdisabled",
	]),
	condition: new Set(["attribute", "entityname", "operator", "value", "valueof", "aggregate"]),
	"link-entity": new Set(["name", "from", "to", "alias", "link-type", "intersect", "visible"]),
};

/** Documented query hint names for the fetch `options` attribute (cc-query-options). */
const DOCUMENTED_QUERY_HINTS = [
	"ForceOrder",
	"DisableRowGoal",
	"EnableOptimizerHotfixes",
	"LoopJoin",
	"MergeJoin",
	"HashJoin",
	"NO_PERFORMANCE_SPOOL",
	"ENABLE_HIST_AMENDMENT_FOR_ASC_KEYS",
];

/** Boolean attribute values accepted by Dataverse ("true"/"1", "false"/"0" — verified). */
function parseBooleanAttr(value: string | null): boolean | undefined {
	if (value === null) return undefined;
	const v = value.trim().toLowerCase();
	if (v === "true" || v === "1") return true;
	if (v === "false" || v === "0") return false;
	return undefined;
}

/**
 * Collect the attributes the tree has no dedicated field for, so they round-trip verbatim (#90):
 * namespace declarations, namespaced attributes, legacy attributes and unknown attributes.
 * Live-validated: Dataverse ignores unknown and namespaced attributes, but rejects a default
 * namespace on <fetch> (0x80041102), so that one is reported as an error and not kept.
 */
function collectElementAttributes(
	el: Element,
	elementName: string,
	issues: ParseIssue[],
): ExtraAttributes | undefined {
	const modelled = MODELLED_ATTRIBUTES[elementName];
	if (!modelled) return undefined;
	const silent = SILENT_LEGACY_ATTRIBUTES[elementName];
	const extras: ExtraAttributes = {};

	for (const attr of Array.from(el.attributes)) {
		const name = attr.name;
		const lower = name.toLowerCase();

		if (lower === "xmlns") {
			issues.push({
				message: `Default namespace (xmlns="${attr.value}") isn't supported by Dataverse; Parse to tree removes it.`,
				element: elementName,
				attribute: name,
				severity: "error",
				kind: "corrected",
			});
			continue;
		}
		if (modelled.has(lower)) continue;

		extras[name] = attr.value;
		if (silent?.has(lower)) continue;

		const isNamespaceDeclaration = attr.namespaceURI === XMLNS_NAMESPACE || lower.startsWith("xmlns:");
		issues.push({
			message: isNamespaceDeclaration
				? `Namespace declaration '${name}' is ignored by Dataverse.`
				: attr.namespaceURI
					? `Namespaced attribute '${name}' is ignored by Dataverse.`
					: `Unknown attribute '${name}' is ignored by Dataverse.`,
			element: elementName,
			attribute: name,
			severity: "info",
			kind: "ignored",
		});
	}

	return Object.keys(extras).length > 0 ? extras : undefined;
}

/** Unknown child element (including namespaced ones): Dataverse rejects it (0x8004111c). */
function reportUnknownChild(child: Element, parentName: string, issues: ParseIssue[]): void {
	issues.push({
		message: `Unknown element <${child.nodeName}> inside <${parentName}> isn't valid FetchXML (Dataverse rejects it); Parse to tree drops it.`,
		element: child.nodeName,
		severity: "error",
		kind: "dropped",
	});
}

/**
 * Aliases Parse to tree corrected in this parse (old → new, or null when removed). References to
 * them elsewhere in the query are updated to match, so the query keeps its meaning (#94).
 */
let linkAliasRenames = new Map<string, string | null>();
let attributeAliasRenames = new Map<string, string | null>();

/**
 * Point references at corrected aliases: order / condition entityname and the valueof qualifier
 * (link-entity aliases), and aggregate order aliases (attribute aliases).
 */
function updateAliasReferences(fetch: FetchNode, issues: ParseIssue[]): void {
	if (linkAliasRenames.size === 0 && attributeAliasRenames.size === 0) return;
	const updated: string[] = [];
	const rename = (map: Map<string, string | null>, value: string | undefined, what: string) => {
		if (value === undefined || !map.has(value)) return value;
		const next = map.get(value) ?? undefined;
		updated.push(`${what} '${value}' → ${next ? `'${next}'` : "removed"}`);
		return next;
	};
	const visitFilter = (filter: FilterNode) => {
		for (const condition of filter.conditions) {
			condition.entityname = rename(linkAliasRenames, condition.entityname, "condition entityname");
			if (condition.valueof) {
				// The corrected alias can itself contain dots ("my.acct.name"): match the longest one
				const valueof = condition.valueof;
				const old = [...linkAliasRenames.keys()]
					.filter((alias) => valueof.startsWith(`${alias}.`))
					.sort((a, b) => b.length - a.length)[0];
				if (old) {
					const qualifier = rename(linkAliasRenames, old, "valueof");
					const attribute = valueof.slice(old.length + 1);
					condition.valueof = qualifier ? `${qualifier}.${attribute}` : attribute;
				}
			}
		}
		filter.subfilters.forEach(visitFilter);
		filter.links?.forEach(visitContainer);
	};
	const visitContainer = (container: EntityNode | LinkEntityNode) => {
		for (const order of container.orders) {
			order.entityname = rename(linkAliasRenames, order.entityname, "order entityname");
			order.alias = rename(attributeAliasRenames, order.alias, "order alias");
		}
		container.filters.forEach(visitFilter);
		container.links.forEach(visitContainer);
	};
	visitContainer(fetch.entity);
	if (updated.length > 0) {
		issues.push({
			message: `Parse to tree also updates the references to the corrected aliases: ${updated.join("; ")}.`,
			element: "fetch",
			severity: "info",
			kind: "corrected",
		});
	}
}

/**
 * Main entry point: Parse FetchXML string to FetchNode
 */
export function parseFetchXml(xmlString: string): ParseResult {
	const issues: ParseIssue[] = [];
	linkAliasRenames = new Map();
	attributeAliasRenames = new Map();

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
				errors: [{ message: `XML parsing error: ${errorText}` }],
				warnings: [],
			};
		}
	} catch (e) {
		return {
			success: false,
			errors: [{ message: `XML parsing error: ${e instanceof Error ? e.message : String(e)}` }],
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
	const extraEntities = fetchElement.querySelectorAll(":scope > entity").length - 1;
	if (extraEntities > 0) {
		issues.push({
			message: `Only one <entity> is allowed; Dataverse uses the first and Parse to tree drops the other ${extraEntities}.`,
			element: "fetch",
			severity: "info",
			kind: "dropped",
		});
	}
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
		// Parse the fetch options and keep any other attributes verbatim
		const options = parseFetchOptions(fetchElement, issues);
		const fetchExtras = collectElementAttributes(fetchElement, "fetch", issues);

		// Parse the entity tree
		const entity = parseEntity(entityElement, issues);

		const fetchNode: FetchNode = {
			id: generateParseId(),
			type: "fetch",
			entity,
			options,
			...(fetchExtras ? { extraAttributes: fetchExtras } : {}),
		};

		updateAliasReferences(fetchNode, issues);
		issues.push(...validateQueryRules(fetchNode));

		const errors: ParseError[] = issues
			.filter((i) => i.severity === "fatal")
			.map((i) => ({ message: i.message }));
		const warnings = issues.filter((i): i is ParseWarning => i.severity !== "fatal");

		return errors.length > 0
			? { success: false, errors, warnings }
			: { success: true, fetchNode, errors, warnings };
	} catch (e) {
		return {
			success: false,
			errors: [{ message: `Parsing error: ${e instanceof Error ? e.message : String(e)}` }],
			warnings: issues.filter((i): i is ParseWarning => i.severity !== "fatal"),
		};
	}
}

/**
 * Parse fetch element attributes into options
 */
function parseFetchOptions(fetchElement: Element, issues: ParseIssue[]): FetchNode["options"] {
	const options: FetchNode["options"] = {};

	// Boolean options. Dataverse accepts true/1 in any case for these (verified), except
	// 'aggregate', which only makes an aggregate query when it is exactly "true".
	const aggregateAttr = fetchElement.getAttribute("aggregate");
	if (aggregateAttr === "true") {
		options.aggregate = true;
	} else if (aggregateAttr !== null && parseBooleanAttr(aggregateAttr) !== false) {
		issues.push({
			message: `aggregate="${aggregateAttr}" doesn't make an aggregate query; Dataverse only recognises "true". Parse to tree treats the query as non-aggregate.`,
			element: "fetch",
			attribute: "aggregate",
			severity: "info",
			kind: "corrected",
		});
	}
	if (parseBooleanAttr(fetchElement.getAttribute("distinct"))) {
		options.distinct = true;
	}
	if (parseBooleanAttr(fetchElement.getAttribute("returntotalrecordcount"))) {
		options.returnTotalRecordCount = true;
	}
	if (parseBooleanAttr(fetchElement.getAttribute("no-lock"))) {
		options.noLock = true;
	}
	if (parseBooleanAttr(fetchElement.getAttribute("latematerialize"))) {
		options.latematerialize = true;
	}
	if (parseBooleanAttr(fetchElement.getAttribute("useraworderby"))) {
		options.useRawOrderBy = true;
	}

	const aggregateLimit = fetchElement.getAttribute("aggregatelimit");
	if (aggregateLimit !== null) {
		const limit = Number(aggregateLimit);
		if (Number.isInteger(limit) && limit >= 1 && limit <= 50000) {
			options.aggregateLimit = limit;
		} else {
			// Dataverse: "Expected value between 1 and 50000 inclusive."
			issues.push({
				message: `'aggregatelimit' must be a whole number from 1 to 50,000 (Dataverse rejects '${aggregateLimit}'); Parse to tree drops it.`,
				element: "fetch",
				attribute: "aggregatelimit",
				severity: "error",
				kind: "dropped",
			});
		}
	}

	const dataSource = fetchElement.getAttribute("datasource");
	if (dataSource !== null) {
		options.dataSource = dataSource;
		// 'retained' is documented; 'bin' is listed as valid in Dataverse's own error message.
		if (dataSource !== "retained" && dataSource !== "bin") {
			issues.push({
				message: `Invalid datasource '${dataSource}'; Dataverse accepts 'retained' (long-term retention) or 'bin'.`,
				element: "fetch",
				attribute: "datasource",
				severity: "error",
				kind: "invalid",
			});
		}
	}

	const optionsAttr = fetchElement.getAttribute("options");
	if (optionsAttr !== null) {
		const hints = optionsAttr
			.split(",")
			.map((h) => h.trim())
			.filter(Boolean);
		if (hints.length > 0) options.queryHints = hints;
		const undocumented = hints.filter(
			(h) => !DOCUMENTED_QUERY_HINTS.some((d) => d.toLowerCase() === h.toLowerCase()),
		);
		if (undocumented.length > 0) {
			// Dataverse accepts (and ignores) unknown hint names, so a typo goes unnoticed
			issues.push({
				message: `Query hint ${undocumented.map((h) => `'${h}'`).join(", ")} isn't a documented option and may have no effect.`,
				element: "fetch",
				attribute: "options",
				severity: "info",
				kind: "advice",
			});
		}
	}

	// Numeric options
	const top = fetchElement.getAttribute("top");
	if (top !== null) {
		const topNum = parseXmlInt(top);
		if (!isNaN(topNum)) {
			options.top = topNum;
		} else {
			issues.push({
				message: `Invalid 'top' value '${top}' (Dataverse rejects it); Parse to tree drops it.`,
				element: "fetch",
				attribute: "top",
				severity: "error",
				kind: "dropped",
			});
		}
	}

	const count = fetchElement.getAttribute("count");
	if (count !== null) {
		const countNum = parseXmlInt(count);
		if (!isNaN(countNum)) {
			options.count = countNum;
		} else {
			issues.push({
				message: `Invalid 'count' value '${count}' (Dataverse rejects it); Parse to tree drops it.`,
				element: "fetch",
				attribute: "count",
				severity: "error",
				kind: "dropped",
			});
		}
	}

	const page = fetchElement.getAttribute("page");
	if (page !== null) {
		const pageNum = parseXmlInt(page);
		if (!isNaN(pageNum)) {
			options.page = pageNum;
		} else {
			issues.push({
				message: `Invalid 'page' value '${page}' (Dataverse rejects it); Parse to tree drops it.`,
				element: "fetch",
				attribute: "page",
				severity: "error",
				kind: "dropped",
			});
		}
	}

	const utcOffset = fetchElement.getAttribute("utc-offset");
	if (utcOffset !== null) {
		const utcNum = parseXmlInt(utcOffset);
		if (!isNaN(utcNum)) {
			options.utcOffset = utcNum;
		} else {
			issues.push({
				message: `Invalid 'utc-offset' value '${utcOffset}' is ignored by Dataverse; Parse to tree drops it.`,
				element: "fetch",
				attribute: "utc-offset",
				severity: "info",
				kind: "ignored",
			});
		}
	}

	// top vs paging and the 5,000-row limit are query-level rules (fetchxmlRules.ts)

	// String options
	const pagingCookie = fetchElement.getAttribute("paging-cookie");
	if (pagingCookie) {
		options.pagingCookie = pagingCookie;
	}


	return options;
}

/**
 * Parse entity element
 */
function parseEntity(entityElement: Element, issues: ParseIssue[]): EntityNode {
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
	// Report prefiltering is documented as enableprefiltering="1"
	if (parseBooleanAttr(entityElement.getAttribute("enableprefiltering"))) {
		entity.enablePrefiltering = true;
	}
	const prefilterParam = entityElement.getAttribute("prefilterparametername");
	if (prefilterParam) {
		entity.prefilterParameterName = prefilterParam;
	}

	const extraAttributes = collectElementAttributes(entityElement, "entity", issues);
	if (extraAttributes) entity.extraAttributes = extraAttributes;

	// Parse child elements
	for (const child of entityElement.children) {
		const tagName = child.nodeName.toLowerCase();

		switch (tagName) {
			case "all-attributes":
				entity.allAttributes = parseAllAttributes(child, issues);
				break;
			case "attribute":
				entity.attributes.push(parseAttribute(child, issues));
				break;
			case "order":
				{
					const order = parseOrder(child, issues);
					if (order) entity.orders.push(order);
				}
				break;
			case "filter":
				entity.filters.push(parseFilter(child, issues));
				break;
			case "link-entity":
				entity.links.push(parseLinkEntity(child, issues));
				break;
			default:
				reportUnknownChild(child, "entity", issues);
		}
	}

	return entity;
}

/**
 * Parse all-attributes element
 * Note: Parameters reserved for future use (potential warnings for deprecated usage)
 */
function parseAllAttributes(_element: Element, _issues: ParseIssue[]): AllAttributesNode {
	return {
		id: generateParseId(),
		type: "all-attributes",
		enabled: true,
	};
}

/**
 * Parse attribute element
 */
function parseAttribute(attrElement: Element, issues: ParseIssue[]): AttributeNode {
	const name = attrElement.getAttribute("name") || "";

	if (!name) {
		issues.push({
			message: "Attribute element missing 'name' (Dataverse rejects it).",
			element: "attribute",
			severity: "error",
			kind: "invalid",
		});
	}

	const extraAttributes = collectElementAttributes(attrElement, "attribute", issues);

	const attr: AttributeNode = {
		id: generateParseId(),
		type: "attribute",
		name,
		...(extraAttributes ? { extraAttributes } : {}),
	};

	// Optional attributes
	const alias = attrElement.getAttribute("alias");
	if (alias) {
		if (!isValidAlias(alias)) {
			const sanitized = sanitizeAlias(alias);
			issues.push({
				message: sanitized
					? `Alias '${alias}' has invalid characters; Parse to tree changes it to '${sanitized}'.`
					: `Alias '${alias}' has no valid characters; Parse to tree removes it.`,
				element: "attribute",
				attribute: "alias",
				severity: "error",
				kind: "corrected",
			});
			attributeAliasRenames.set(alias, sanitized || null);
			if (sanitized) attr.alias = sanitized;
		} else {
			attr.alias = alias;
		}
	}

	if (parseBooleanAttr(attrElement.getAttribute("groupby"))) {
		attr.groupby = true;
	}

	const aggregate = attrElement.getAttribute("aggregate");
	if (aggregate) {
		const validAggregates = ["sum", "count", "countcolumn", "min", "max", "avg"];
		if (validAggregates.includes(aggregate)) {
			attr.aggregate = aggregate as AttributeNode["aggregate"];
		} else if (aggregate === "rowaggregate") {
			// Dataverse: 0x8004111a "The value 'rowaggregate' is not a valid aggregate operation"
			issues.push({
				message: `'rowaggregate' isn't an aggregate function; to count child records use rowaggregate="CountChildren" instead. Parse to tree drops it.`,
				element: "attribute",
				attribute: "aggregate",
				severity: "error",
				kind: "dropped",
			});
		} else {
			issues.push({
				message: `Invalid aggregate value '${aggregate}'; Parse to tree drops it.`,
				element: "attribute",
				attribute: "aggregate",
				severity: "error",
				kind: "dropped",
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
			issues.push({
				message: `Invalid dategrouping value '${dategrouping}'; Parse to tree drops it.`,
				element: "attribute",
				attribute: "dategrouping",
				severity: "error",
				kind: "dropped",
			});
		}
	}

	// "1"/"0" mean true/false here too (verified: same groupings as "true"/"false")
	const usertimezone = parseBooleanAttr(attrElement.getAttribute("usertimezone"));
	if (usertimezone !== undefined) {
		attr.usertimezone = usertimezone;
	}

	if (parseBooleanAttr(attrElement.getAttribute("distinct"))) {
		attr.distinct = true;
	}

	const rowAggregate = attrElement.getAttribute("rowaggregate");
	if (rowAggregate !== null) {
		attr.rowAggregate = rowAggregate;
		if (rowAggregate.toLowerCase() !== "countchildren") {
			issues.push({
				message: `rowaggregate '${rowAggregate}' isn't documented (only 'CountChildren' is); Dataverse may reject it.`,
				element: "attribute",
				attribute: "rowaggregate",
				severity: "warning",
				kind: "invalid",
			});
		}
	}

	return attr;
}

/**
 * Parse order element
 */
function parseOrder(orderElement: Element, issues: ParseIssue[]): OrderNode | null {
	const attribute = orderElement.getAttribute("attribute") || undefined;
	const alias = orderElement.getAttribute("alias") || undefined;

	if (!attribute && !alias) {
		issues.push({
			message: "Order element has neither 'attribute' nor 'alias' (Dataverse rejects it); Parse to tree drops it.",
			element: "order",
			severity: "error",
			kind: "dropped",
		});
		return null;
	}

	const extraAttributes = collectElementAttributes(orderElement, "order", issues);
	const order: OrderNode = {
		id: generateParseId(),
		type: "order",
		...(attribute ? { attribute } : {}),
		...(alias ? { alias } : {}),
		...(extraAttributes ? { extraAttributes } : {}),
	};

	// Dataverse accepts exactly 'true', 'false', '1' or '0' here (0x80041119 otherwise)
	const descending = orderElement.getAttribute("descending");
	if (descending === "true" || descending === "1") {
		order.descending = true;
	} else if (descending !== null && descending !== "false" && descending !== "0") {
		issues.push({
			message: `Invalid descending value '${descending}'; Dataverse only accepts 'true', 'false', '1' or '0' (0x80041119). Parse to tree sorts ascending.`,
			element: "order",
			attribute: "descending",
			severity: "error",
			kind: "corrected",
		});
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
function parseFilter(filterElement: Element, issues: ParseIssue[]): FilterNode {
	const type = filterElement.getAttribute("type") || "and";
	if (type !== "and" && type !== "or") {
		issues.push({
			message: `Invalid filter type '${type}' (Dataverse rejects it); Parse to tree uses 'and'.`,
			element: "filter",
			attribute: "type",
			severity: "error",
			kind: "corrected",
		});
	}
	const extraAttributes = collectElementAttributes(filterElement, "filter", issues);

	const filter: FilterNode = {
		id: generateParseId(),
		type: "filter",
		conjunction: type === "or" ? "or" : "and",
		conditions: [],
		subfilters: [],
		links: [],
	};

	// The only valid hint is "union" (case-sensitive; anything else is rejected: 0x80041122)
	const hint = filterElement.getAttribute("hint");
	if (hint === "union") {
		filter.hint = "union";
	} else if (hint !== null) {
		issues.push({
			message: `Invalid filter hint '${hint}'; the only valid value is 'union' (Dataverse rejects it: 0x80041122). Parse to tree drops it.`,
			element: "filter",
			attribute: "hint",
			severity: "error",
			kind: "dropped",
		});
	}
	if (extraAttributes) filter.extraAttributes = extraAttributes;
	if (parseBooleanAttr(filterElement.getAttribute("isquickfindfields"))) {
		filter.isQuickFindFields = true;
	}
	if (parseBooleanAttr(filterElement.getAttribute("overridequickfindrecordlimitenabled"))) {
		filter.overrideQuickFindRecordLimitEnabled = true;
	}
	if (parseBooleanAttr(filterElement.getAttribute("overridequickfindrecordlimitdisabled"))) {
		filter.overrideQuickFindRecordLimitDisabled = true;
	}

	// Parse child elements
	for (const child of filterElement.children) {
		const tagName = child.nodeName.toLowerCase();

		switch (tagName) {
			case "condition":
				filter.conditions.push(parseCondition(child, issues));
				break;
			case "filter":
				filter.subfilters.push(parseFilter(child, issues));
				break;
			case "link-entity":
				// Parse link-entity inside filter (for any/not any/all/not all)
				filter.links.push(parseLinkEntity(child, issues));
				break;
			default:
				reportUnknownChild(child, "filter", issues);
		}
	}

	return filter;
}

/**
 * Parse condition element
 */
function parseCondition(condElement: Element, issues: ParseIssue[]): ConditionNode {
	const attribute = condElement.getAttribute("attribute") || "";
	const operatorAttr = condElement.getAttribute("operator");
	const operator = operatorAttr || "eq";

	if (!attribute) {
		issues.push({
			message: "Condition element missing 'attribute' (Dataverse rejects it).",
			element: "condition",
			severity: "error",
			kind: "invalid",
		});
	}
	if (!operatorAttr) {
		issues.push({
			message: "Condition element missing 'operator' (Dataverse rejects it); Parse to tree uses 'eq'.",
			element: "condition",
			attribute: "operator",
			severity: "error",
			kind: "corrected",
		});
	} else if (operatorAttr === "neq") {
		issues.push({
			message: "Operator 'neq' is deprecated; it works the same as 'ne'.",
			element: "condition",
			attribute: "operator",
			severity: "info",
			kind: "advice",
		});
	} else if (!KNOWN_OPERATORS.has(operatorAttr)) {
		issues.push({
			message: `Operator '${operatorAttr}' isn't a documented condition operator; Dataverse may reject it.`,
			element: "condition",
			attribute: "operator",
			severity: "warning",
			kind: "invalid",
		});
	}

	const extraAttributes = collectElementAttributes(condElement, "condition", issues);

	const condition: ConditionNode = {
		id: generateParseId(),
		type: "condition",
		attribute,
		operator: operator as OperatorType,
		...(extraAttributes ? { extraAttributes } : {}),
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
		const valueAttributes: Record<string, Record<string, string>> = {};
		for (const valEl of valueElements) {
			// Not trimmed: the value is kept as written, spaces included (#181)
			const valText = valEl.textContent ?? "";
			// Try to parse as number if it looks like one
			const numVal = parseFloat(valText);
			if (!isNaN(numVal) && valText === numVal.toString()) {
				values.push(numVal);
			} else {
				values.push(valText);
			}
			// uiname / uitype on lookup values in saved views, kept verbatim (#181)
			if (valEl.attributes.length > 0) {
				valueAttributes[valText] = Object.fromEntries(Array.from(valEl.attributes, (a) => [a.name, a.value]));
			}
		}
		condition.value = values;
		if (Object.keys(valueAttributes).length > 0) condition.valueAttributes = valueAttributes;
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
function parseLinkEntity(linkElement: Element, issues: ParseIssue[]): LinkEntityNode {
	const name = linkElement.getAttribute("name") || "";
	const from = linkElement.getAttribute("from") || "";
	const to = linkElement.getAttribute("to") || "";

	if (!name) {
		issues.push({
			message: "Link-entity element missing 'name' (Dataverse rejects it).",
			element: "link-entity",
			severity: "error",
			kind: "invalid",
		});
	}

	// Live-validated: from="" and to="" together return HTTP 200 with zero rows — silently wrong.
	const hasFrom = linkElement.hasAttribute("from");
	const hasTo = linkElement.hasAttribute("to");
	if (hasFrom && hasTo && !from && !to) {
		issues.push({
			message: `Link-entity '${name}' has empty 'from' and 'to' attributes; Dataverse returns no rows for this. Set both, or remove them for a system many-to-many relationship.`,
			element: "link-entity",
			severity: "fatal",
			kind: "invalid",
		});
	} else if (!from || !to) {
		const missing = [!from && "'from'", !to && "'to'"].filter(Boolean).join(" and ");
		issues.push({
			message: `Link-entity '${name}' is missing ${missing}. Dataverse can only infer the join for a system many-to-many relationship.`,
			element: "link-entity",
			severity: "warning",
			kind: "invalid",
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
		// Undocumented, but accepted by Dataverse (listed in its own link-type error message;
		// verified against a live org).
		"natural",
	];

	let linkType: LinkType = "inner";
	if (validLinkTypes.includes(linkTypeAttr as LinkType)) {
		linkType = linkTypeAttr as LinkType;
	} else {
		issues.push({
			message: `Invalid link-type '${linkTypeAttr}' (Dataverse rejects it); Parse to tree uses 'inner'.`,
			element: "link-entity",
			attribute: "link-type",
			severity: "error",
			kind: "corrected",
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
	if (alias) {
		if (!isValidAlias(alias)) {
			const sanitized = sanitizeAlias(alias);
			issues.push({
				message: sanitized
					? `Alias '${alias}' has invalid characters; Parse to tree changes it to '${sanitized}'.`
					: `Alias '${alias}' has no valid characters; Parse to tree removes it.`,
				element: "link-entity",
				attribute: "alias",
				severity: "info",
				kind: "corrected",
			});
			linkAliasRenames.set(alias, sanitized || null);
			if (sanitized) link.alias = sanitized;
		} else {
			link.alias = alias;
		}
	}

	const extraAttributes = collectElementAttributes(linkElement, "link-entity", issues);
	if (extraAttributes) link.extraAttributes = extraAttributes;

	if (parseBooleanAttr(linkElement.getAttribute("intersect"))) {
		link.intersect = true;
	}

	const visible = parseBooleanAttr(linkElement.getAttribute("visible"));
	if (visible !== undefined) {
		link.visible = visible;
	}

	// Parse child elements
	for (const child of linkElement.children) {
		const tagName = child.nodeName.toLowerCase();

		switch (tagName) {
			case "all-attributes":
				link.allAttributes = parseAllAttributes(child, issues);
				break;
			case "attribute":
				link.attributes.push(parseAttribute(child, issues));
				break;
			case "order":
				{
					const order = parseOrder(child, issues);
					if (order) link.orders.push(order);
				}
				break;
			case "filter":
				link.filters.push(parseFilter(child, issues));
				break;
			case "link-entity":
				link.links.push(parseLinkEntity(child, issues));
				break;
			default:
				reportUnknownChild(child, "link-entity", issues);
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
