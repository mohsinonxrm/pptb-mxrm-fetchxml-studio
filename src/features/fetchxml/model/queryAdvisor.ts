/**
 * Query Advisor (#97): documented performance anti-patterns and query limits, reported before the
 * query runs. Called by validateQueryRules, so findings reach Validate, Parse to Tree and the
 * property editors through the same path as the other query rules.
 *
 * Principle (#89): advice never blocks. Performance guidance is severity "info", kind "advice", so
 * Execute never stops or asks for it. Only the documented hard limits, which Dataverse rejects with
 * a documented error code, are errors; a limit whose outcome can't be determined client-side is a
 * warning ("Execute anyway"). Every check cites the Microsoft Learn article it comes from.
 *
 * Rules that need attribute metadata are in adviseQueryMetadata: orders on choice attributes,
 * conditions on large text attributes (Memo, or String with MaxLength > 850), conditions on
 * calculated or formula attributes (SourceType), and paging without a unique order (primary key).
 * The Advisories tab runs them with the cached metadata; validateQueryRules (Validate, Parse to
 * Tree, the tree and the property editors) has no metadata and doesn't. When an entity's metadata
 * isn't loaded, its rules are skipped: they never guess.
 */

import { generateFetchXml } from "./fetchxml";
import type { RuleFinding } from "./fetchxmlRules";
import type { ConditionNode, EntityNode, FetchNode, FilterNode, LinkEntityNode, OrderNode } from "./nodes";

/** Link types that return no attributes (orders there are already an error, see fetchxmlRules). */
const NON_RETURNING = new Set(["exists", "in", "any", "not any", "all", "not all"]);

/** https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/join-tables#limitations */
export const MAX_LINK_ENTITIES = 15;
/**
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows#condition-limits
 * The docs count condition and link-entity elements together; live (2026-10-17) Dataverse counts
 * conditions only: 500 conditions + a link-entity run, 501 conditions fail with 0x8004430C.
 */
export const MAX_CONDITIONS = 500;
/** https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results#fetchxml-paging-models */
export const SIMPLE_PAGING_LIMIT = 50_000;
/** Default page size for standard entities (page-results) */
const DEFAULT_PAGE_SIZE = 5000;
/** https://learn.microsoft.com/power-apps/developer/data-platform/webapi/compose-http-requests-handle-errors#web-api-url-and-versions */
export const MAX_URL_LENGTH = 32_768;

type Container = EntityNode | LinkEntityNode;

interface Walk {
	/** Every entity / link-entity, root first */
	containers: { container: Container; isRoot: boolean; inFilter: boolean }[];
	conditions: ConditionNode[];
	/** Orders with the container they're in */
	orders: { order: OrderNode; container: Container; isRoot: boolean }[];
}

function walk(fetch: FetchNode): Walk {
	const result: Walk = { containers: [], conditions: [], orders: [] };
	const visit = (container: Container, isRoot: boolean, inFilter: boolean) => {
		result.containers.push({ container, isRoot, inFilter });
		container.orders.forEach((order) => result.orders.push({ order, container, isRoot }));
		const visitFilter = (filter: FilterNode) => {
			result.conditions.push(...filter.conditions);
			filter.subfilters.forEach(visitFilter);
			filter.links?.forEach((l) => visit(l, false, true));
		};
		container.filters.forEach(visitFilter);
		container.links.forEach((l) => visit(l, false, inFilter));
	};
	visit(fetch.entity, true, false);
	return result;
}

/** Operators that match with wildcards (wildcard-characters) */
const PATTERN_OPERATORS = new Set(["like", "not-like", "begins-with", "not-begin-with"]);
/** Operators with an implicit leading wildcard */
const ENDS_WITH_OPERATORS = new Set(["ends-with", "not-end-with"]);

/**
 * Whether a pattern can't use an index: it starts with a wildcard (%, _ or [), or with hyphens or
 * apostrophes followed by a wildcard ("-%", "'%751"), which collation rules make behave the same.
 * begins-with values get an implicit trailing % (so begins-with "-" is "-%"). A bracket escape
 * ([%], [_], [[]) is a literal character, not a wildcard (wildcard-characters, #98).
 * https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns#avoid-leading-wild-cards-in-filter-conditions
 */
export function hasLeadingWildcard(operator: string, value: unknown): boolean {
	if (!PATTERN_OPERATORS.has(operator) || typeof value !== "string" || value === "") return false;
	const pattern = (operator === "begins-with" || operator === "not-begin-with" ? `${value}%` : value).replace(/^[-']*/, "");
	return /^[%_[]/.test(pattern) && !/^\[[%_[]\]/.test(pattern);
}

/**
 * Leading-wildcard advice for one condition (#97). Also used for a results-grid column filter's
 * condition (#98), which isn't in the tree yet.
 * https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns#avoid-leading-wild-cards-in-filter-conditions
 * https://learn.microsoft.com/power-apps/developer/data-platform/wildcard-characters#dont-use-leading-wild-cards
 */
export function adviseCondition(condition: ConditionNode): RuleFinding[] {
	if (condition.valueof) return [];
	const slow = "so Dataverse can't use an index: the query can be slow and may be throttled (0x80048644)";
	const base = { element: "condition", nodeId: condition.id, severity: "info", kind: "advice", rule: "leading-wildcard", target: condition.attribute } as const;
	if (ENDS_WITH_OPERATORS.has(condition.operator)) {
		return [{ ...base, attribute: "operator",
			message: `The '${condition.operator}' operator works like a leading wildcard, ${slow}. Use begins-with where you can, or Dataverse search.` }];
	}
	if (hasLeadingWildcard(condition.operator, condition.value)) {
		const value = String(condition.value);
		const start = /^[-']/.test(value) ? "starts with a hyphen or apostrophe before a wildcard, which works like a leading wildcard," : "starts with a wildcard,";
		return [{ ...base, attribute: "value",
			message: `The value '${value}' ${start} ${slow}. Start the value with literal text, or use Dataverse search.` }];
	}
	return [];
}

export function adviseQuery(fetch: FetchNode): RuleFinding[] {
	const findings: RuleFinding[] = [];
	const push = (f: RuleFinding) => findings.push(f);
	const { containers, conditions, orders } = walk(fetch);
	const options = fetch.options;
	const aggregate = options.aggregate === true;
	const links = containers.flatMap(({ container, inFilter }) => (container.type === "link-entity" ? [{ link: container, inFilter }] : []));

	// ---- leading wildcards: Dataverse can't use an index and throttles these queries (0x80048644)
	for (const condition of conditions) adviseCondition(condition).forEach(push);

	// ---- columns: "We strongly discourage returning all columns" (select-columns; query-antipatterns
	// LargeAmountOfAttributes). With no attribute and no all-attributes, the root returns all columns.
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/select-columns
	if (!aggregate) {
		for (const { container } of containers) {
			if (!container.allAttributes?.enabled) continue;
			push({ element: "all-attributes", nodeId: container.allAttributes.id, severity: "info", kind: "advice", rule: "all-attributes",
				message: `All attributes of ${container.name} are returned, which makes the query slower and can cause timeouts. Select only the attributes you need.` });
		}
		const root = fetch.entity;
		if (root.attributes.length === 0 && !root.allAttributes?.enabled) {
			push({ element: "entity", nodeId: root.id, severity: "info", kind: "advice", rule: "no-attributes",
				message: `No attributes of ${root.name} are selected, so Dataverse returns all of them, which makes the query slower and can cause timeouts. Select only the attributes you need.` });
		}
	}

	// ---- orders on link-entity attributes: slower (query-antipatterns OrderOnPropertiesFromJoinedTables),
	// and Dataverse might not return a paging cookie, so paging falls back to simple paging (page-results)
	// https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns#avoid-ordering-by-columns-in-related-tables
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results#queries-that-dont-support-paging-cookies
	if (!aggregate) {
		const linkKeys = new Set(links.flatMap(({ link }) => [link.alias, link.name].filter((k): k is string => !!k)));
		for (const { order, container, isRoot } of orders) {
			const inLink = container.type === "link-entity" && !NON_RETURNING.has(container.linkType);
			const byEntityname = isRoot && !!order.entityname && order.entityname !== fetch.entity.name && linkKeys.has(order.entityname);
			if (!inLink && !byEntityname) continue;
			push({ element: "order", attribute: byEntityname ? "entityname" : "attribute", nodeId: order.id, severity: "info", kind: "advice", rule: "order-link", target: order.attribute,
				message: `Sorting by a link-entity's attribute makes the query slower, and Dataverse might not return a paging cookie (paging then stops at ${SIMPLE_PAGING_LIMIT.toLocaleString("en-US")} records). Sort by the root entity's attributes where you can.` });
		}
	}

	// ---- distinct without an order: Dataverse can't add its default primary key order
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/overview#return-distinct-results
	if (options.distinct && !aggregate && orders.length === 0) {
		push({ element: "fetch", attribute: "distinct", nodeId: fetch.id, severity: "info", kind: "advice", rule: "distinct-order",
			message: `With distinct, add at least one order: Dataverse can't add its default primary key order, so pages can be inconsistent.` });
	}

	// ---- simple paging stops at 50,000 records
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results#simple-paging
	if (options.page !== undefined && options.page >= 1 && !options.pagingCookie && options.top === undefined) {
		const size = options.count ?? DEFAULT_PAGE_SIZE;
		if (options.page * size > SIMPLE_PAGING_LIMIT) {
			push({ element: "fetch", attribute: "page", nodeId: fetch.id, severity: "info", kind: "advice", rule: "simple-paging-limit",
				message: `Without a paging cookie, paging can't go past the first ${SIMPLE_PAGING_LIMIT.toLocaleString("en-US")} records; page ${options.page} with ${size} records per page goes beyond that. Use the paging cookie from the previous page.` });
		}
	}

	// ---- aggregate record limit (8004E023) when nothing narrows the records
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/aggregate-data#aggregate-query-record-limit-error
	if (aggregate && options.aggregateLimit === undefined && conditions.length === 0 && !links.some((l) => l.inFilter)) {
		push({ element: "fetch", attribute: "aggregate", nodeId: fetch.id, severity: "info", kind: "advice", rule: "aggregate-limit",
			message: `Aggregate queries fail when more than ${SIMPLE_PAGING_LIMIT.toLocaleString("en-US")} records match (0x8004E023), and this query has no filter. For a large entity, add a filter or set an aggregate row limit.` });
	}

	// ---- legacy fetch options. (Documented query hints get no finding (#90); the editor already says
	// to use them only when Microsoft support recommends it.)
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/optimize-performance#no-lock
	if (options.noLock) {
		push({ element: "fetch", attribute: "no-lock", nodeId: fetch.id, severity: "info", kind: "advice", rule: "no-lock",
			message: `no-lock is a legacy setting and no longer necessary; you can remove it.` });
	}

	// ---- limits
	// More than 15 link-entities: 0x8004430D (join-tables). Verified live 2026-10-17: 16 joined
	// link-entities fail, and link-entities inside a filter don't count (14 joined + 2 any run).
	const joined = links.filter((l) => !l.inFilter);
	if (joined.length > MAX_LINK_ENTITIES) {
		for (const { link } of joined.slice(MAX_LINK_ENTITIES)) {
			push({ element: "link-entity", nodeId: link.id, severity: "error", kind: "invalid", rule: "link-entity-limit",
				message: `A query can have at most ${MAX_LINK_ENTITIES} link-entities; this one has ${joined.length} (Dataverse rejects it: 0x8004430D). Remove link-entities, or split the query.` });
		}
	}
	// More than 500 conditions: 0x8004430C (filter-rows › Condition limits; link-entities don't count, verified live)
	if (conditions.length > MAX_CONDITIONS) {
		push({ element: "fetch", nodeId: fetch.id, severity: "error", kind: "invalid", rule: "condition-limit",
			message: `A query can have at most ${MAX_CONDITIONS} conditions; this one has ${conditions.length} (Dataverse rejects it: 0x8004430C). Combine conditions with the in operator.` });
	}
	// GET URLs are limited to 32,768 characters (0x80060888). Studio sends FetchXML in the URL; the
	// host and organization URL add to it, so the exact limit can't be determined here.
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/retrieve-data#use-fetchxml-within-a-batch-request
	const encodedLength = encodeURIComponent(generateFetchXml(fetch)).length;
	if (encodedLength > MAX_URL_LENGTH) {
		push({ element: "fetch", nodeId: fetch.id, severity: "warning", kind: "invalid", rule: "url-length",
			message: `The URL-encoded FetchXML is ${encodedLength.toLocaleString("en-US")} characters, but Web API URLs are limited to ${MAX_URL_LENGTH.toLocaleString("en-US")}, so Dataverse may reject it (0x80060888). Shorten the query, for example with fewer attributes or conditions.` });
	}

	return findings;
}

// ---------------------------------------------------------------------------------------------
// Rules that need attribute metadata (#97)
// ---------------------------------------------------------------------------------------------

/** What the metadata rules need to know about an attribute (AttributeMetadata subset). */
export interface AdvisorAttributeInfo {
	/** AttributeTypeName.Value, e.g. "PicklistType" (as toAttributeTypeInfo gives it) */
	typeName: string;
	/** AttributeMetadata.SourceType: 0 simple, 1 calculated, 2 rollup, 3 formula, 4 prompt */
	sourceType?: number;
	/** StringAttributeMetadata / MemoAttributeMetadata.MaxLength */
	maxLength?: number;
}

/** What the metadata rules know about one entity; a missing part means "not loaded". */
export interface AdvisorEntityInfo {
	primaryIdAttribute?: string;
	attributes?: ReadonlyMap<string, AdvisorAttributeInfo>;
}

/** Metadata per entity logical name, from the metadata cache */
export type AdvisorMetadata = ReadonlyMap<string, AdvisorEntityInfo>;

/**
 * Choice attribute types. Dataverse sorts them by the localized label, not the stored number
 * (order-rows › Choice columns). Yes/No (Boolean) attributes are two-option choices. Verified live
 * (2026-10-18): ValidateFetchXmlExpression reports OrderOnEnumAttribute (TypeCode 4) for Picklist,
 * State, Status and Boolean orders, and nothing with useraworderby. Multi-select choices can't be
 * sorted at all (attributeCapabilities: an error in the order editor).
 */
const CHOICE_TYPES = new Set(["PicklistType", "StateType", "StatusType", "BooleanType"]);

/**
 * "All memo columns or string columns with a MaxLength greater than 850 are defined in Dataverse as
 * large text columns." https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns#avoid-using-conditions-on-large-text-columns
 */
export const LARGE_TEXT_MAX_LENGTH = 850;

/**
 * Formula and calculated column values "are calculated in real-time when they're retrieved"
 * (query-antipatterns); rollup values are stored, so filtering on them isn't this anti-pattern.
 * SourceType values: https://learn.microsoft.com/power-apps/developer/data-platform/specialized-columns#common-elements-and-characteristics
 */
const CALCULATED_SOURCE_TYPES: ReadonlyMap<number, string> = new Map([
	[1, "calculated"],
	[3, "formula"],
]);

/** Why an attribute is a large text attribute, or undefined when it isn't (or its length is unknown). */
export function largeTextReason(info: AdvisorAttributeInfo | undefined): string | undefined {
	if (!info) return undefined;
	if (info.typeName === "MemoType") return "multiline text (Memo)";
	if (info.typeName === "StringType" && info.maxLength !== undefined && info.maxLength > LARGE_TEXT_MAX_LENGTH) {
		return `text with a maximum length of ${info.maxLength.toLocaleString("en-US")}`;
	}
	return undefined;
}

/** Link-entities by alias, or by name when the name is unique, for entityname lookups. */
function linkLookup(fetch: FetchNode): (key: string | undefined) => LinkEntityNode | undefined {
	const links = walk(fetch).containers.flatMap(({ container }) => (container.type === "link-entity" ? [container] : []));
	return (key) => {
		if (!key) return undefined;
		const byAlias = links.find((l) => l.alias === key);
		if (byAlias) return byAlias;
		const byName = links.filter((l) => l.name === key);
		return byName.length === 1 ? byName[0] : undefined;
	};
}

/**
 * Advice that needs attribute metadata (#97). Pure. An entity whose metadata isn't in `metadata`
 * (not loaded yet, or the request failed) is skipped, so nothing is guessed.
 */
/** What the advisor knows about the query's results (#170) */
export interface AdvisorContext {
	/** This query's results are paged: more than one page loaded or available (Load more, Retrieve All) */
	resultsPaged?: boolean;
}

export function adviseQueryMetadata(fetch: FetchNode, metadata: AdvisorMetadata, context: AdvisorContext = {}): RuleFinding[] {
	const findings: RuleFinding[] = [];
	const push = (f: RuleFinding) => findings.push(f);
	const { orders } = walk(fetch);
	const options = fetch.options;
	const aggregate = options.aggregate === true;
	const findLink = linkLookup(fetch);
	const infoOf = (entity: string, attribute: string | undefined) =>
		attribute ? metadata.get(entity)?.attributes?.get(attribute) : undefined;

	// ---- orders on choice attributes: sorted by the localized label, which needs more compute
	// https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns#avoid-ordering-by-choice-columns
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows#choice-columns
	// Aggregate queries order by alias. useraworderby sorts choices by their number instead.
	if (!aggregate && !options.useRawOrderBy) {
		for (const { order, container, isRoot } of orders) {
			if (!order.attribute) continue;
			const target = isRoot && order.entityname && order.entityname !== fetch.entity.name ? findLink(order.entityname) : container;
			if (!target) continue;
			const info = infoOf(target.name, order.attribute);
			if (!info || !CHOICE_TYPES.has(info.typeName)) continue;
			const kind = info.typeName === "BooleanType" ? "Yes/No" : "choice";
			push({ element: "order", attribute: "attribute", nodeId: order.id, severity: "info", kind: "advice", rule: "order-choice", target: order.attribute,
				message: `Sorting by the ${kind} attribute '${order.attribute}' sorts by its label in the user's language, which makes the query slower. Sort by another attribute where you can, or turn on "Sort option sets by value" (useraworderby) on the fetch.` });
		}
	}

	// ---- conditions on large text attributes, and on calculated or formula attributes
	// https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns#avoid-using-conditions-on-large-text-columns
	// https://learn.microsoft.com/power-apps/developer/data-platform/query-antipatterns#avoid-using-formula-or-calculated-columns-in-filter-conditions
	// The docs apply to any condition on these attributes, whatever the operator. Verified live
	// (2026-10-18): ValidateFetchXmlExpression reports eq, not-null, like, not-like, begins-with and
	// ends-with on a Memo attribute (TypeCode 3), and conditions on calculated attributes (TypeCode
	// 8), but nothing for rollup attributes.
	const visitContainer = (container: Container) => {
		const visitFilter = (filter: FilterNode) => {
			for (const condition of filter.conditions) {
				const target = condition.entityname ? findLink(condition.entityname) : container;
				if (!target) continue;
				const info = infoOf(target.name, condition.attribute);
				const base = { element: "condition", attribute: "attribute", nodeId: condition.id, severity: "info", kind: "advice", target: condition.attribute } as const;
				const large = largeTextReason(info);
				if (large) {
					push({ ...base, rule: "large-text",
						message: `'${condition.attribute}' is a large text attribute (${large}), too large to index, so conditions on it make the query slow. Use Dataverse search to search text like this.` });
				}
				const computed = info?.sourceType !== undefined ? CALCULATED_SOURCE_TYPES.get(info.sourceType) : undefined;
				if (computed) {
					push({ ...base, rule: "calculated-filter",
						message: `'${condition.attribute}' is a ${computed} attribute: Dataverse calculates it for every record to apply the condition, so the query is slow and may be throttled (0x80048744). Filter on the attributes it's calculated from where you can.` });
				}
			}
			filter.subfilters.forEach(visitFilter);
			filter.links?.forEach(visitContainer);
		};
		container.filters.forEach(visitFilter);
		container.links.forEach(visitContainer);
	};
	visitContainer(fetch.entity);

	// ---- paging without a unique order
	// "If you need consistent orders and paging, you must include some unique values" (order-rows ›
	// Deterministic ordering is important); "queries should order on the primary key for the table"
	// (› Best practices for orders when paging data). With no order, Dataverse adds a primary key
	// order (documented; verified live 2026-10-18: rows come back by contactid and the paging cookie
	// holds contactid), so that gets no advice. Distinct queries have their own rule, and adding the
	// key would defeat distinct.
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/order-rows#ordering-and-paging
	// https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results
	// Kept as advice (#170): live, Dataverse broke ties by the primary key in every case tried, but
	// that isn't documented. Paged means the query sets count or page, or its results came in more
	// than one page: FetchXML Studio pages any query past 5,000 rows (Load more, Retrieve All).
	const root = fetch.entity;
	const primaryId = metadata.get(root.name)?.primaryIdAttribute;
	const pagedByQuery = options.count !== undefined || options.page !== undefined;
	const paged = pagedByQuery || context.resultsPaged === true;
	if (primaryId && paged && !aggregate && !options.distinct && options.top === undefined && orders.length > 0) {
		const unique = root.orders.some((o) => o.attribute === primaryId && (!o.entityname || o.entityname === root.name));
		if (!unique) {
			// Dataverse applies link-entity orders after the entity's orders (order-rows › Process
			// link-entity orders first), so a key added to the entity wouldn't be the last order
			const linkOrders = orders.some(({ container }) => container.type === "link-entity");
			const problem = `${pagedByQuery ? "The query is paged" : "The results are paged"}, but no order is on a unique attribute such as ${primaryId}, so records can repeat or go missing between pages.`;
			push({ element: "entity", nodeId: root.id, severity: "info", kind: "advice", rule: "paging-order", target: primaryId,
				message: linkOrders
					? `${problem} Dataverse sorts by link-entity orders after the entity's orders: to end with ${primaryId}, first move them to the entity with entityname.`
					: `${problem} Add ${primaryId} as the last order.`,
				...(linkOrders ? {} : { fix: { kind: "add-root-order", attribute: primaryId } as const }) });
		}
	}

	return findings;
}
