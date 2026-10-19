/**
 * Query-level FetchXML rules that need the whole tree (#90b). Shared by the parser (Validate,
 * Execute, Parse to Tree) and the property editors, so both report the same thing.
 *
 * Principle (#89): what the FetchXML reference documents is supported, plus what is verified to
 * work; only what Dataverse is known to reject is an error. Every rule below was executed against a
 * live Dataverse org through the PPTB request path; the error code Dataverse returns is quoted in
 * each message. Rules that need attribute metadata live in attributeCapabilities.ts.
 */

import type {
	AttributeNode,
	ConditionNode,
	EntityNode,
	FetchNode,
	FilterNode,
	LinkEntityNode,
	OrderNode,
} from "./nodes";
import { operatorRequiresValue, operatorValueShape } from "./operators";

export type RuleSeverity = "error" | "warning" | "info";
export type RuleKind = "invalid" | "ignored" | "advice";

export interface RuleFinding {
	message: string;
	element: string;
	attribute?: string;
	severity: RuleSeverity;
	kind: RuleKind;
	/** Tree node the finding is about, so editors can show it on that node */
	nodeId: string;
}

type Container = EntityNode | LinkEntityNode;

interface Scope {
	container: Container;
	isRoot: boolean;
	/** Nesting level of filters, starting at 1 for a container's own filter */
	filters: { filter: FilterNode; depth: number }[];
}

/** Every entity / link-entity in the query, including link-entities inside filters. */
function collectScopes(fetch: FetchNode): Scope[] {
	const scopes: Scope[] = [];
	const visit = (container: Container, isRoot: boolean) => {
		const scope: Scope = { container, isRoot, filters: [] };
		scopes.push(scope);
		const visitFilter = (filter: FilterNode, depth: number) => {
			scope.filters.push({ filter, depth });
			filter.subfilters.forEach((f) => visitFilter(f, depth + 1));
			filter.links?.forEach((l) => visit(l, false));
		};
		container.filters.forEach((f) => visitFilter(f, 1));
		container.links.forEach((l) => visit(l, false));
	};
	visit(fetch.entity, true);
	return scopes;
}

const isLink = (c: Container): c is LinkEntityNode => c.type === "link-entity";

/** Link types allowed for a link-entity inside a filter (others, or none: 0x8004111c, verified live). */
export const FILTER_LINK_TYPES: ReadonlySet<string> = new Set(["any", "not any", "all", "not all"]);
/** Link types that return no attributes: attribute, all-attributes and order fail with 0x80041121 (verified live). */
export const NON_RETURNING_LINK_TYPES: ReadonlySet<string> = new Set(["exists", "in", "any", "not any", "all", "not all"]);

/** Operators a `valueof` comparison accepts (filter-rows: "Limitations on column comparison filters"). */
export const VALUEOF_OPERATORS: ReadonlySet<string> = new Set(["eq", "ne", "gt", "ge", "lt", "le"]);

/** What a condition's `valueof` can refer to with "qualifier.attribute". */
export interface ValueOfScope {
	/** The root entity's logical name: always reachable (verified, though not documented) */
	rootName: string;
	/** Aliases of the joined link-entities it can reach */
	aliases: string[];
	/** Alias of the joined link-entity whose filter holds the condition (its own record) */
	selfAlias?: string;
	/** Joined aliases it can't reach because they're joined after (or inside) its link-entity */
	laterAliases: string[];
}

/**
 * `valueof` reach per condition, verified live (2026-10-15). The docs show a root filter comparing to
 * a link-entity alias; Dataverse also accepts these, and rejects anything else with 0x80040216:
 * - conditions in the root entity's filters (including entityname conditions, and any / not any /
 *   all / not all link-entities there): every joined link-entity alias, at any depth;
 * - conditions in a joined link-entity's filter: its own alias and link-entities joined before it in
 *   the query (not its own link-entities, nor later ones);
 * - a link-entity inside a filter has the reach of the filter's entity / link-entity; its own alias,
 *   and aliases of other link-entities inside filters, are never reachable;
 * - the root entity's logical name works everywhere.
 */
export function valueofScopes(fetch: FetchNode): Map<string, ValueOfScope> {
	const joined: LinkEntityNode[] = [];
	const collect = (c: Container) =>
		c.links.forEach((l) => {
			joined.push(l);
			collect(l);
		});
	collect(fetch.entity);
	const all = joined.flatMap((l) => (l.alias ? [l.alias] : []));
	const map = new Map<string, ValueOfScope>();
	const visitFilters = (filters: FilterNode[], scope: ValueOfScope) => {
		const visit = (f: FilterNode) => {
			f.conditions.forEach((c) => map.set(c.id, scope));
			f.subfilters.forEach(visit);
			// A link-entity inside a filter keeps the reach of the scope that holds the filter
			f.links?.forEach((l) => visitContainer(l, scope));
		};
		filters.forEach(visit);
	};
	const visitContainer = (c: Container, scope: ValueOfScope) => {
		visitFilters(c.filters, scope);
		c.links.forEach((l) => {
			const index = joined.indexOf(l);
			const reachable = joined.slice(0, index + 1).flatMap((x) => (x.alias ? [x.alias] : []));
			const linkScope: ValueOfScope =
				index === -1
					? scope
					: {
							rootName: fetch.entity.name,
							aliases: reachable,
							selfAlias: l.alias,
							laterAliases: all.filter((a) => !reachable.includes(a)),
						};
			visitContainer(l, linkScope);
		});
	};
	visitContainer(fetch.entity, { rootName: fetch.entity.name, aliases: all, laterAliases: [] });
	return map;
}

/**
 * Attributes an aggregate `<order alias>` can see. Dataverse resolves the alias only within the
 * order's own scope: the entity / link-entity the order is in, or, for an order in the root entity
 * with `entityname`, that link-entity (verified: other aliases fail with 0x8004110a).
 */
function orderAliasScope(order: OrderNode, scope: Scope, links: LinkEntityNode[]): AttributeNode[] {
	if (scope.isRoot && order.entityname) {
		const target =
			links.find((l) => l.alias === order.entityname) ?? links.find((l) => l.name === order.entityname);
		if (target) return target.attributes;
	}
	return scope.container.attributes;
}

export function validateQueryRules(fetch: FetchNode): RuleFinding[] {
	const findings: RuleFinding[] = [];
	const aggregate = fetch.options.aggregate === true;
	const scopes = collectScopes(fetch);
	const links = scopes.map((s) => s.container).filter(isLink);
	const push = (f: RuleFinding) => findings.push(f);

	// ---- attributes
	for (const scope of scopes) {
		for (const attr of scope.container.attributes) {
			const name = attr.name || "(unnamed)";
			const base = { element: "attribute", nodeId: attr.id } as const;

			if (attr.rowAggregate !== undefined) {
				if (aggregate) {
					// SQL error 4104 even with groupby; 0x8004111e without
					push({ ...base, attribute: "rowaggregate", severity: "error", kind: "invalid",
						message: `Count child records (rowaggregate) doesn't work in an aggregate query; Dataverse rejects '${name}'.` });
				} else if (!scope.isRoot) {
					push({ ...base, attribute: "rowaggregate", severity: "error", kind: "invalid",
						message: `Count child records (rowaggregate) only works on the root entity; inside a link-entity Dataverse fails with an internal error.` });
				} else if (!attr.alias) {
					push({ ...base, attribute: "rowaggregate", severity: "error", kind: "invalid",
						message: `Count child records (rowaggregate) on '${name}' needs an alias; without one Dataverse fails with a SQL error.` });
				}
			}

			if (aggregate) {
				if (attr.groupby && attr.aggregate) {
					push({ ...base, attribute: "groupby", severity: "error", kind: "invalid",
						message: `'${name}' can't be both Group by and an aggregate (Dataverse rejects it: 0x80041137).` });
				} else if (!attr.groupby && !attr.aggregate && attr.rowAggregate === undefined) {
					push({ ...base, severity: "error", kind: "invalid",
						message: `In an aggregate query every attribute needs Group by or an aggregate function; Dataverse rejects '${name}' (0x8004111e).` });
				}
				if (!attr.alias) {
					push({ ...base, attribute: "alias", severity: "error", kind: "invalid",
						message: `In an aggregate query every attribute needs an alias; '${name}' has none (Dataverse rejects it: 0x8004110b).` });
				}
				if (attr.dategrouping && attr.aggregate && !attr.groupby) {
					push({ ...base, attribute: "dategrouping", severity: "info", kind: "ignored",
						message: `Date grouping on '${name}' is ignored: it only applies to Group by.` });
				}
				if (attr.distinct && attr.aggregate !== "countcolumn") {
					push({ ...base, attribute: "distinct", severity: "info", kind: "advice",
						message: `'distinct' on '${name}' is only documented with the countcolumn aggregate.` });
				}
			} else {
				if (attr.aggregate) {
					push({ ...base, attribute: "aggregate", severity: "error", kind: "invalid",
						message: `'${name}' has aggregate="${attr.aggregate}", but the query isn't an aggregate query (Dataverse rejects it: 0x8004111a).` });
				}
				if (attr.groupby) {
					push({ ...base, attribute: "groupby", severity: "error", kind: "invalid",
						message: `'${name}' has Group by, but the query isn't an aggregate query (Dataverse rejects it: 0x8004112e).` });
				}
				const ignored = [
					attr.dategrouping && "dategrouping",
					attr.usertimezone !== undefined && "usertimezone",
					attr.distinct && "distinct",
				].filter((x): x is string => typeof x === "string");
				if (ignored.length > 0) {
					push({ ...base, attribute: ignored[0], severity: "info", kind: "ignored",
						message: `${ignored.map((a) => `'${a}'`).join(", ")} on '${name}' ${ignored.length > 1 ? "are" : "is"} ignored outside aggregate queries.` });
				}
			}
		}
	}

	// ---- aliases: unique across the whole query, case-insensitive (0x80041130). An alias also
	// clashes with the auto-generated name of an unaliased root attribute (its logical name).
	const rootNames = new Set(
		fetch.entity.attributes.filter((a) => !a.alias && a.name).map((a) => a.name.toLowerCase()),
	);
	const seenAliases = new Map<string, AttributeNode>();
	for (const scope of scopes) {
		for (const attr of scope.container.attributes) {
			if (!attr.alias) continue;
			const key = attr.alias.toLowerCase();
			if (seenAliases.has(key)) {
				push({ element: "attribute", attribute: "alias", nodeId: attr.id, severity: "error", kind: "invalid",
					message: `Alias '${attr.alias}' is used more than once (aliases are case-insensitive); Dataverse rejects it (0x80041130).` });
			} else if (rootNames.has(key) && !(scope.isRoot && attr.name.toLowerCase() === key)) {
				push({ element: "attribute", attribute: "alias", nodeId: attr.id, severity: "error", kind: "invalid",
					message: `Alias '${attr.alias}' clashes with the attribute '${key}' selected without an alias; Dataverse rejects it (0x80041130).` });
			}
			seenAliases.set(key, attr);
		}
	}
	const seenLinkAliases = new Set<string>();
	for (const link of links) {
		if (!link.alias) continue;
		if (seenLinkAliases.has(link.alias)) {
			push({ element: "link-entity", attribute: "alias", nodeId: link.id, severity: "error", kind: "invalid",
				message: `Link-entity alias '${link.alias}' is used more than once; Dataverse rejects it (0x80041130).` });
		}
		seenLinkAliases.add(link.alias);
	}
	// Link types by placement and what they return (verified live, 2026-10-14). Joined link-entities
	// accept all nine types, including the four the docs describe for use inside a filter.
	const filterLinkIds = new Set(scopes.flatMap((s) => s.filters.flatMap(({ filter }) => (filter.links ?? []).map((l) => l.id))));
	for (const link of links) {
		const base = { element: "link-entity", nodeId: link.id } as const;
		if (filterLinkIds.has(link.id) && !FILTER_LINK_TYPES.has(link.linkType)) {
			push({ ...base, attribute: "link-type", severity: "error", kind: "invalid",
				message: `A link-entity inside a filter must be any, not any, all or not all (Dataverse rejects '${link.linkType}': 0x8004111c).` });
		}
		if (NON_RETURNING_LINK_TYPES.has(link.linkType)) {
			const returns = [
				link.attributes.length > 0 && "attributes",
				link.allAttributes?.enabled && "all-attributes",
				link.orders.length > 0 && "orders",
			].filter(Boolean);
			if (returns.length > 0) {
				push({ ...base, attribute: "link-type", severity: "error", kind: "invalid",
					message: `A '${link.linkType}' link-entity returns no attributes, so it can't have ${returns.join(" or ")} (Dataverse rejects it: 0x80041121).` });
			}
		}
		if (link.visible !== undefined) {
			push({ ...base, attribute: "visible", severity: "info", kind: "ignored",
				message: `The visible attribute isn't in the FetchXML reference, and Dataverse ignores it: the results are the same with any value.` });
		}
	}

	// ---- orders
	const entityNames = new Set<string>([fetch.entity.name, ...links.flatMap((l) => [l.name, l.alias ?? ""])]);
	let anyOrder = false;
	for (const scope of scopes) {
		for (const order of scope.container.orders) {
			anyOrder = true;
			const base = { element: "order", nodeId: order.id } as const;
			if (aggregate && order.attribute) {
				push({ ...base, attribute: "attribute", severity: "error", kind: "invalid",
					message: `In an aggregate query, order by the attribute's alias instead of attribute '${order.attribute}' (Dataverse rejects it: 0x80041131).` });
			} else if (!aggregate && order.alias) {
				push({ ...base, attribute: "alias", severity: "error", kind: "invalid",
					message: `Ordering by alias '${order.alias}' only works in aggregate queries; use 'attribute' instead (Dataverse rejects it: 0x80041132).` });
			} else if (aggregate && order.alias) {
				const visible = orderAliasScope(order, scope, links);
				if (!visible.some((a) => a.alias === order.alias)) {
					const where = scope.isRoot && order.entityname ? `link-entity '${order.entityname}'` : isLink(scope.container) ? `link-entity '${scope.container.alias ?? scope.container.name}'` : "the root entity";
					push({ ...base, attribute: "alias", severity: "error", kind: "invalid",
						message: `Alias '${order.alias}' isn't defined on ${where}; an aggregate order can only use aliases from its own entity (Dataverse rejects it: 0x8004110a).` });
				}
			}
			if (order.entityname) {
				if (!scope.isRoot) {
					push({ ...base, attribute: "entityname", severity: "error", kind: "invalid",
						message: `'entityname' isn't allowed on an order inside a link-entity (Dataverse rejects it: 0x80041108). Move the order to the root entity.` });
				} else if (!entityNames.has(order.entityname)) {
					push({ ...base, attribute: "entityname", severity: "error", kind: "invalid",
						message: `No link-entity has the alias or name '${order.entityname}' (Dataverse rejects it: 0x80041108).` });
				}
			}
		}
	}

	// ---- fetch options
	const paging = [
		fetch.options.count !== undefined && "count",
		fetch.options.page !== undefined && "page",
		fetch.options.returnTotalRecordCount && "returntotalrecordcount",
	].filter((x): x is string => typeof x === "string");
	if (fetch.options.top !== undefined && paging.length > 0) {
		push({ element: "fetch", attribute: "top", nodeId: fetch.id, severity: "error", kind: "invalid",
			message: `'top' can't be used with ${paging.map((p) => `'${p}'`).join(", ")} (Dataverse rejects it: 0x80040203). Use either top or paging.` });
	}
	for (const [name, value] of [
		["top", fetch.options.top],
		["count", fetch.options.count],
	] as const) {
		if (value !== undefined && value > 5000) {
			push({ element: "fetch", attribute: name, nodeId: fetch.id, severity: "info", kind: "advice",
				message: `'${name}' is ${value}, but Dataverse returns at most 5,000 rows per request.` });
		}
	}
	if (fetch.options.page !== undefined && fetch.options.page < 1) {
		push({ element: "fetch", attribute: "page", nodeId: fetch.id, severity: "error", kind: "invalid",
			message: `Page number must be 1 or more (Dataverse rejects page ${fetch.options.page}: 0x80040203).` });
	}
	if (fetch.options.aggregateLimit !== undefined) {
		if (!aggregate) {
			push({ element: "fetch", attribute: "aggregatelimit", nodeId: fetch.id, severity: "info", kind: "ignored",
				message: `'aggregatelimit' is ignored: the query isn't an aggregate query.` });
		} else if (anyOrder) {
			// Undocumented platform behaviour, reproduced on every table tried
			push({ element: "fetch", attribute: "aggregatelimit", nodeId: fetch.id, severity: "error", kind: "invalid",
				message: `Dataverse returns a SQL error when an aggregate query with 'aggregatelimit' also has an order. Remove the order or the limit.` });
		}
	}

	// ---- filters
	const reach = valueofScopes(fetch);
	const unionFilters: { filter: FilterNode; depth: number }[] = [];
	for (const scope of scopes) {
		for (const { filter, depth } of scope.filters) {
			const base = { element: "filter", nodeId: filter.id } as const;
			if (filter.hint === "union") {
				unionFilters.push({ filter, depth });
				if (filter.conjunction !== "or") {
					push({ ...base, attribute: "hint", severity: "error", kind: "invalid",
						message: `The union hint only works on an 'or' filter (Dataverse rejects it: 0x80041122).` });
				} else if (depth > 3) {
					push({ ...base, attribute: "hint", severity: "info", kind: "ignored",
						message: `The union hint is more than three filter levels deep, so Dataverse ignores it.` });
				}
			}
			if (filter.overrideQuickFindRecordLimitEnabled && filter.overrideQuickFindRecordLimitDisabled) {
				push({ ...base, attribute: "overridequickfindrecordlimitenabled", severity: "info", kind: "advice",
					message: `Both quick find record limit overrides are set; which one applies isn't documented. Keep one.` });
			}
			// valueof (verified live): only eq/ne/gt/ge/lt/le compare two attributes (others: 0x80048415);
			// operators without a value ignore it; "x.attr" must be reachable (valueofScopes; 0x80040216).
			for (const condition of filter.conditions) {
				if (condition.valueof === undefined || condition.valueof === "") continue;
				const cbase = { element: "condition", attribute: "valueof", nodeId: condition.id } as const;
				if (!operatorRequiresValue(condition.operator)) {
					push({ ...cbase, severity: "info", kind: "ignored",
						message: `Dataverse ignores valueof here: the '${condition.operator}' operator doesn't compare to a value.` });
					continue;
				}
				if (!VALUEOF_OPERATORS.has(condition.operator)) {
					push({ ...cbase, severity: "error", kind: "invalid",
						message: `Comparing to another attribute only works with eq, ne, gt, ge, lt or le (Dataverse rejects '${condition.operator}': 0x80048415).` });
				}
				const dot = condition.valueof.indexOf(".");
				const qualifier = dot > 0 ? condition.valueof.slice(0, dot) : undefined;
				const scope = reach.get(condition.id);
				if (qualifier && scope && qualifier !== scope.rootName && !scope.aliases.includes(qualifier)) {
					push({ ...cbase, severity: "error", kind: "invalid",
						message: scope.laterAliases.includes(qualifier)
							? `'${qualifier}' is joined after this link-entity (or inside it), so this filter can't refer to it. Compare from the root entity's filter instead (Dataverse rejects it: 0x80040216).`
							: `'${qualifier}' isn't the alias of a joined link-entity. To compare to another entity's attribute, give its link-entity an alias and use alias.attribute (Dataverse rejects it: 0x80040216).` });
				}
			}
			// Value shape per operator (verified live 2026-10-15, #87). A condition the user is still
			// editing can have no attribute or operator yet (picking an attribute clears the operator):
			// say what's missing, as the parser does, instead of judging the value of an unknown operator.
			for (const condition of filter.conditions) {
				if (!condition.attribute) {
					push({ element: "condition", attribute: "attribute", nodeId: condition.id, severity: "error", kind: "invalid",
						message: "Choose an attribute: a condition needs one (Dataverse rejects it without one)." });
				}
				if (!condition.operator) {
					push({ element: "condition", attribute: "operator", nodeId: condition.id, severity: "error", kind: "invalid",
						message: "Choose an operator: a condition needs one (Dataverse rejects it without one)." });
					continue;
				}
				const issue = conditionValueIssue(condition);
				if (issue) {
					push({ element: "condition", attribute: "value", nodeId: condition.id, severity: issue.severity, kind: issue.severity === "error" ? "invalid" : "ignored",
						message: issue.message });
				}
			}
			// condition/@aggregate isn't in the condition reference. Verified live: Dataverse runs the
			// condition as a plain row condition (the value is converted to the attribute's type), so
			// FetchXML has no HAVING and removing it doesn't change the results.
			for (const condition of filter.conditions) {
				if (condition.aggregate) {
					push({ element: "condition", attribute: "aggregate", nodeId: condition.id, severity: "info", kind: "ignored",
						message: `Dataverse ignores aggregate="${condition.aggregate}" on a condition: it compares each record's ${condition.attribute} value, not an aggregate result. FetchXML can't filter on aggregate values.` });
				}
			}
		}
	}
	for (const { filter } of unionFilters.slice(1)) {
		push({ element: "filter", attribute: "hint", nodeId: filter.id, severity: "error", kind: "invalid",
			message: `Only one filter per query can use the union hint (Dataverse rejects it: 0x80041122).` });
	}
	// The reference allows one <filter> per entity / link-entity; Dataverse accepts more and combines
	// them with 'and' (verified live), so this is advice only (#94).
	for (const scope of scopes) {
		for (const filter of scope.container.filters.slice(1)) {
			push({ element: "filter", nodeId: filter.id, severity: "info", kind: "advice",
				message: `${scope.container.name} has more than one filter; the FetchXML reference allows one. Dataverse combines them with 'and' — consider nesting them in one filter.` });
		}
	}

	return findings;
}

const isWholeNumber = (v: unknown): boolean =>
	(typeof v === "number" && Number.isInteger(v)) || (typeof v === "string" && /^-?\d+$/.test(v.trim()));
const isFourDigitYear = (v: unknown): boolean => isWholeNumber(v) && Number(v) >= 1000 && Number(v) <= 9999;
const isEmptyValue = (v: unknown): boolean => v === undefined || v === null || v === "";

/**
 * Whether a condition's value has the shape its operator needs. Verified live (2026-10-15):
 * - single-value operators need the value attribute (<value> elements: 0x8004111e);
 * - "x" operators need a whole number above 0 (0x80041139 not a number, 0x80040203 zero or less);
 * - in-fiscal-year needs a four-digit year (0x80040203);
 * - between and the fiscal period-and-year operators need exactly two values (0x80041106);
 * - in / not-in / contain-values need at least one <value> (0x80040203);
 * - operators without a value ignore one.
 */
export function conditionValueIssue(condition: ConditionNode): { severity: RuleSeverity; message: string } | undefined {
	const op = condition.operator;
	const value = condition.value;
	const values = Array.isArray(value) ? value : undefined;
	switch (operatorValueShape(op)) {
		case "none":
			return isEmptyValue(value) || (values && values.length === 0)
				? undefined
				: { severity: "info", message: `Dataverse ignores the value: the '${op}' operator doesn't take one.` };
		case "single":
			if (condition.valueof) return undefined;
			if (values) return { severity: "error", message: `The '${op}' operator takes one value in the value attribute, not <value> elements (Dataverse rejects it: 0x8004111e).` };
			return isEmptyValue(value) ? { severity: "error", message: `The '${op}' operator needs a value (Dataverse rejects it without one: 0x8004111e).` } : undefined;
		case "integer":
			if (values || isEmptyValue(value)) return { severity: "error", message: `The '${op}' operator needs a whole number above 0 in the value attribute (0x8004111e).` };
			if (!isWholeNumber(value)) return { severity: "error", message: `The '${op}' operator needs a whole number, not '${String(value)}' (Dataverse rejects it: 0x80041139).` };
			return Number(value) < 1 ? { severity: "error", message: `The '${op}' operator needs a number above 0 (Dataverse rejects it: 0x80040203).` } : undefined;
		case "year":
			if (values || isEmptyValue(value)) return { severity: "error", message: `The '${op}' operator needs a fiscal year in the value attribute (0x8004111e).` };
			return isFourDigitYear(value) ? undefined : { severity: "error", message: `The '${op}' operator needs a four-digit fiscal year, not '${String(value)}' (Dataverse rejects it: 0x80040203).` };
		case "two":
			return values && values.length === 2 && values.every((v) => !isEmptyValue(v))
				? undefined
				: { severity: "error", message: `The '${op}' operator needs exactly two <value> elements (Dataverse rejects it: 0x80041106).` };
		case "fiscalPeriodYear":
			if (!values || values.length !== 2) return { severity: "error", message: `The '${op}' operator needs two <value> elements, the period and the year (Dataverse rejects it: 0x80041106).` };
			if (!isWholeNumber(values[0]) || Number(values[0]) < 1) return { severity: "error", message: `The fiscal period for '${op}' must be a whole number from 1.` };
			return isFourDigitYear(values[1]) ? undefined : { severity: "error", message: `The fiscal year for '${op}' must be a four-digit year.` };
		case "multiple":
			return values && values.length > 0
				? undefined
				: { severity: "error", message: `The '${op}' operator needs at least one <value> element (Dataverse rejects a value attribute or no value: 0x80040203).` };
	}
}
