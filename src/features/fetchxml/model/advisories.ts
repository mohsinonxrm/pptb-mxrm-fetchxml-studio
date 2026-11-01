/**
 * The Advisories tab's list (#97): the query rules' findings (validateQueryRules, including the
 * Query Advisor in queryAdvisor.ts), the metadata rules (adviseQueryMetadata), and what Dataverse's
 * ValidateFetchXmlExpression reports when the user checks with Dataverse. Pure: no host or React.
 *
 * Advice never blocks Execute (#89): this only describes findings. Settings Dataverse ignores
 * (kind "ignored") aren't advisories; the property editors and Validate show them.
 */

import { validateQueryRules, type AdvisorFix, type AdvisorRuleId, type RuleFinding } from "./fetchxmlRules";
import { adviseQueryMetadata, type AdvisorContext, type AdvisorMetadata } from "./queryAdvisor";
import type { ConditionNode, EntityNode, FetchNode, FilterNode, LinkEntityNode, NodeId, OrderNode } from "./nodes";

export type AdvisoryLevel = "error" | "warning" | "advice";
/** Who reported it: FetchXML Studio, Dataverse, or both ("Dataverse agrees") */
export type AdvisorySource = "fxs" | "dataverse" | "both";

export interface Advisory {
	/** Unique in the list, for React keys */
	id: string;
	level: AdvisoryLevel;
	title: string;
	detail: string;
	/** The tree node it's about; the location selects it */
	nodeId?: NodeId;
	/** Where it is, e.g. "contact › filter › fullname like" */
	location: string;
	doc: { label: string; href: string };
	source: AdvisorySource;
	/** Dataverse's own message, when Dataverse reported it */
	dataverseText?: string;
	fix?: AdvisorFix;
	rule?: AdvisorRuleId;
	/** The attribute it's about, matched with Dataverse's messages */
	target?: string;
	/** The entity `target` belongs to, matched with Dataverse's messages (nodeEntities) */
	entity?: AdvisoryEntity;
}

/** The entity an attribute belongs to: the root, or a link-entity (by name and alias) */
export interface AdvisoryEntity {
	name: string;
	alias?: string;
	isRoot: boolean;
}

const LEARN = "https://learn.microsoft.com/power-apps/developer/data-platform";
const ANTIPATTERNS = `${LEARN}/query-antipatterns`;
const FETCHXML = `${LEARN}/fetchxml`;

const ANTIPATTERNS_LABEL = "Query anti-patterns";

/** Title and Learn page per Query Advisor rule */
const RULES: Record<AdvisorRuleId, { title: string; doc: { label: string; href: string } }> = {
	"leading-wildcard": { title: "Leading wildcard", doc: { label: ANTIPATTERNS_LABEL, href: `${ANTIPATTERNS}#avoid-leading-wild-cards-in-filter-conditions` } },
	"all-attributes": { title: "All attributes returned", doc: { label: "Select columns", href: `${FETCHXML}/select-columns` } },
	"no-attributes": { title: "No attributes selected", doc: { label: "Select columns", href: `${FETCHXML}/select-columns` } },
	"order-link": { title: "Order on a link-entity attribute", doc: { label: "Page results", href: `${FETCHXML}/page-results#queries-that-dont-support-paging-cookies` } },
	"distinct-order": { title: "Distinct without an order", doc: { label: "Return distinct results", href: `${FETCHXML}/overview#return-distinct-results` } },
	"simple-paging-limit": { title: "Paging past 50,000 records", doc: { label: "Page results", href: `${FETCHXML}/page-results#simple-paging` } },
	"aggregate-limit": { title: "Aggregate query without a filter", doc: { label: "Aggregate data", href: `${FETCHXML}/aggregate-data#aggregate-query-record-limit-error` } },
	"no-lock": { title: "Legacy no-lock setting", doc: { label: "Optimize performance", href: `${FETCHXML}/optimize-performance#no-lock` } },
	"link-entity-limit": { title: "Too many link-entities", doc: { label: "Join tables", href: `${FETCHXML}/join-tables#limitations` } },
	"condition-limit": { title: "Too many conditions", doc: { label: "Filter rows", href: `${FETCHXML}/filter-rows#condition-limits` } },
	"url-length": { title: "FetchXML too long for a URL", doc: { label: "Retrieve data", href: `${FETCHXML}/retrieve-data#use-fetchxml-within-a-batch-request` } },
	"order-choice": { title: "Order on a choice attribute", doc: { label: ANTIPATTERNS_LABEL, href: `${ANTIPATTERNS}#avoid-ordering-by-choice-columns` } },
	"large-text": { title: "Condition on a large text attribute", doc: { label: ANTIPATTERNS_LABEL, href: `${ANTIPATTERNS}#avoid-using-conditions-on-large-text-columns` } },
	"calculated-filter": { title: "Condition on a calculated attribute", doc: { label: ANTIPATTERNS_LABEL, href: `${ANTIPATTERNS}#avoid-using-formula-or-calculated-columns-in-filter-conditions` } },
	"paging-order": { title: "Paging without a unique order", doc: { label: "Order rows", href: `${FETCHXML}/order-rows#ordering-and-paging` } },
};

const LEVEL_TITLE: Record<AdvisoryLevel, string> = {
	error: "Dataverse rejects this",
	warning: "Dataverse might reject this",
	advice: "Query advice",
};

const LEVEL_ORDER: Record<AdvisoryLevel, number> = { error: 0, warning: 1, advice: 2 };

export const LEVEL_LABELS: Record<AdvisoryLevel, string> = {
	error: "Will be rejected",
	warning: "Might be rejected",
	advice: "Performance advice",
};

function levelOf(f: RuleFinding): AdvisoryLevel | undefined {
	if (f.severity === "error") return "error";
	if (f.severity === "warning") return "warning";
	return f.kind === "advice" ? "advice" : undefined;
}

/** The FetchXML reference page for an element (rules without their own page) */
const referenceDoc = (element: string) => ({ label: "FetchXML reference", href: `${FETCHXML}/reference/${element}` });

// ---- locations

type Container = EntityNode | LinkEntityNode;

const containerLabel = (c: Container) => (c.type === "link-entity" && c.alias ? `${c.name} (${c.alias})` : c.name);
const qualified = (entityname: string | undefined, attribute: string | undefined) =>
	`${entityname ? `${entityname}.` : ""}${attribute ?? ""}`;
const conditionLabel = (c: ConditionNode) => `${qualified(c.entityname, c.attribute)} ${c.operator}`;
const orderLabel = (o: OrderNode) => `order ${qualified(o.entityname, o.attribute ?? o.alias)}`;

/** A location per node: its entity / link-entity path, then the node (consecutive filters shown once). */
export function nodeLocations(fetch: FetchNode): Map<NodeId, string> {
	const map = new Map<NodeId, string>([[fetch.id, "fetch"]]);
	const join = (path: string[]) => path.join(" › ");
	const visitContainer = (c: Container, parent: string[]) => {
		const path = [...parent, containerLabel(c)];
		map.set(c.id, join(path));
		if (c.allAttributes) map.set(c.allAttributes.id, join([...path, "all attributes"]));
		c.attributes.forEach((a) => map.set(a.id, join([...path, a.name])));
		c.orders.forEach((o) => map.set(o.id, join([...path, orderLabel(o)])));
		const visitFilter = (f: FilterNode, filterPath: string[]) => {
			map.set(f.id, join(filterPath));
			f.conditions.forEach((cond) => map.set(cond.id, join([...filterPath, conditionLabel(cond)])));
			f.subfilters.forEach((sub) => visitFilter(sub, filterPath));
			f.links?.forEach((l) => visitContainer(l, filterPath));
		};
		c.filters.forEach((f) => visitFilter(f, [...path, "filter"]));
		c.links.forEach((l) => visitContainer(l, path));
	};
	visitContainer(fetch.entity, []);
	return map;
}

/**
 * The entity each node's attribute belongs to: its entity / link-entity, or for a condition or a
 * root order with entityname, the link-entity that names (by alias, else by a unique name).
 */
export function nodeEntities(fetch: FetchNode): Map<NodeId, AdvisoryEntity> {
	const map = new Map<NodeId, AdvisoryEntity>();
	const links: LinkEntityNode[] = [];
	const collect = (c: Container) => {
		const visitFilter = (f: FilterNode) => {
			f.subfilters.forEach(visitFilter);
			f.links?.forEach((l) => {
				links.push(l);
				collect(l);
			});
		};
		c.filters.forEach(visitFilter);
		c.links.forEach((l) => {
			links.push(l);
			collect(l);
		});
	};
	collect(fetch.entity);
	const entityOf = (c: Container, isRoot: boolean): AdvisoryEntity => ({ name: c.name, alias: c.type === "link-entity" ? c.alias : undefined, isRoot });
	const root = entityOf(fetch.entity, true);
	const byEntityname = (key: string | undefined, fallback: AdvisoryEntity): AdvisoryEntity => {
		if (!key || key === fetch.entity.name) return fallback;
		const named = links.find((l) => l.alias === key) ?? (links.filter((l) => l.name === key).length === 1 ? links.find((l) => l.name === key) : undefined);
		return named ? entityOf(named, false) : fallback;
	};
	map.set(fetch.id, root);
	const visit = (c: Container, isRoot: boolean) => {
		const own = entityOf(c, isRoot);
		map.set(c.id, own);
		if (c.allAttributes) map.set(c.allAttributes.id, own);
		c.attributes.forEach((a) => map.set(a.id, own));
		c.orders.forEach((o) => map.set(o.id, isRoot ? byEntityname(o.entityname, own) : own));
		const visitFilter = (f: FilterNode) => {
			map.set(f.id, own);
			f.conditions.forEach((cond) => map.set(cond.id, byEntityname(cond.entityname, own)));
			f.subfilters.forEach(visitFilter);
			f.links?.forEach((l) => visit(l, false));
		};
		c.filters.forEach(visitFilter);
		c.links.forEach((l) => visit(l, false));
	};
	visit(fetch.entity, true);
	return map;
}

// ---- FetchXML Studio's findings

/**
 * Every advisory FetchXML Studio finds without a request: the query rules, plus the metadata rules
 * when metadata is given. Sorted by level (stable, so tree order within a level).
 */
export function collectAdvisories(fetch: FetchNode, metadata?: AdvisorMetadata, context?: AdvisorContext): Advisory[] {
	const findings = [...validateQueryRules(fetch), ...(metadata ? adviseQueryMetadata(fetch, metadata, context) : [])];
	const locations = nodeLocations(fetch);
	const entities = nodeEntities(fetch);
	const list: Advisory[] = [];
	findings.forEach((f, index) => {
		const level = levelOf(f);
		if (!level) return;
		const rule = f.rule ? RULES[f.rule] : undefined;
		list.push({
			id: `fxs:${index}:${f.rule ?? f.element}:${f.nodeId}`,
			level,
			title: rule?.title ?? LEVEL_TITLE[level],
			detail: f.message,
			nodeId: f.nodeId,
			location: locations.get(f.nodeId) ?? fetch.entity.name,
			doc: rule?.doc ?? referenceDoc(f.element),
			source: "fxs",
			fix: f.fix,
			rule: f.rule,
			target: f.target,
			entity: entities.get(f.nodeId),
		});
	});
	return list.sort((a, b) => LEVEL_ORDER[a.level] - LEVEL_ORDER[b.level]);
}

// ---- Dataverse's findings (ValidateFetchXmlExpression)

/**
 * A ValidatorIssue from ValidateFetchXmlExpression.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/validatorissue
 */
export interface DataverseIssue {
	LocalizedMessageText: string;
	Severity: number;
	TypeCode: number;
	OptionalPropertyBag?: unknown;
}

/**
 * TypeCodes and the rules they match. TypeCode isn't documented; these are the codes Dataverse
 * returned on a live org (2026-10-18), with the anti-pattern each message describes:
 * 2 leading wildcard ("Leading wild card on like/notlike conditions… Attribute fullname"),
 * 3 large text column ("Like Condition on Attribute description will not perform well…", for any
 *   operator), 4 order on a choice ("Ordering on Enum variables…"), 5 order on a joined table
 *   ("Ordering on columns from joined tables…"), 6 too many select attributes (no attribute;
 *   AttributeCount / AttributeLimit), 8 condition on a calculated column ("Filters on calculated
 *   columns…"). Any other code is shown as Dataverse wrote it.
 */
export const DATAVERSE_TYPE_RULES: Readonly<Record<number, readonly AdvisorRuleId[]>> = {
	2: ["leading-wildcard"],
	3: ["large-text"],
	4: ["order-choice"],
	5: ["order-link"],
	6: ["all-attributes", "no-attributes"],
	8: ["calculated-filter"],
};

/** Codes whose message isn't about one attribute */
const QUERY_LEVEL_TYPES = new Set([6]);

/**
 * The attribute a Dataverse message is about: the OptionalPropertyBag's "Attribute" value
 * ({ Count, Keys: [...], Values: [...] } as returned live), else "Attribute <name>" in the text.
 */
export function issueAttribute(issue: DataverseIssue): string | undefined {
	const bag = issue.OptionalPropertyBag as { Keys?: unknown; Values?: unknown; Attribute?: unknown } | undefined;
	if (bag && Array.isArray(bag.Keys) && Array.isArray(bag.Values)) {
		const i = bag.Keys.indexOf("Attribute");
		if (i >= 0 && typeof bag.Values[i] === "string" && bag.Values[i]) return bag.Values[i] as string;
	}
	if (bag && typeof bag.Attribute === "string" && bag.Attribute) return bag.Attribute;
	const match = /\bAttribute\s+([A-Za-z_]\w*)\b(?!\s*:)/.exec(issue.LocalizedMessageText ?? "");
	return match?.[1];
}

/** Bag keys that might name the entity. None were returned live (see issueEntity). */
const ENTITY_KEYS = ["Entity", "EntityName", "EntityAlias", "Alias", "Table"];

/**
 * The entity or alias a Dataverse message is about, if it says. Verified live (2026-10-18): it
 * doesn't. Every TypeCode returned only Attribute (plus ConditionOperator, Reason, AttributeCount /
 * AttributeLimit), and the text names only the attribute. So matching usually falls back to the
 * attribute, preferring the root entity's finding (mergeDataverseCheck). The keys are read in case a
 * later version adds them.
 */
export function issueEntity(issue: DataverseIssue): string | undefined {
	const bag = issue.OptionalPropertyBag as { Keys?: unknown; Values?: unknown } & Record<string, unknown> | undefined;
	if (!bag) return undefined;
	for (const key of ENTITY_KEYS) {
		if (Array.isArray(bag.Keys) && Array.isArray(bag.Values)) {
			const i = bag.Keys.indexOf(key);
			if (i >= 0 && typeof bag.Values[i] === "string" && bag.Values[i]) return bag.Values[i] as string;
		}
		if (typeof bag[key] === "string" && bag[key]) return bag[key] as string;
	}
	return undefined;
}

/**
 * The finding a Dataverse message agrees with: same rule and attribute (any, for query-level codes),
 * not matched yet. When the message names an entity, only findings on that entity (name or alias);
 * otherwise the root entity's finding first. Live, Dataverse only checks the root entity's
 * conditions (a leading wildcard inside a link-entity's filter isn't reported, one with entityname
 * in the root's filter is), and its messages don't say which entity, so "Attribute name" most
 * likely means the root's. Ties go to the first in tree order.
 */
function findAgreeing(list: Advisory[], rules: readonly AdvisorRuleId[], attribute: string | undefined, entity: string | undefined, queryLevel: boolean): Advisory | undefined {
	const want = attribute?.toLowerCase();
	const candidates = list.filter(
		(a) => a.source === "fxs" && a.rule !== undefined && rules.includes(a.rule) && (queryLevel || a.target?.toLowerCase() === want),
	);
	if (queryLevel) return candidates[0];
	if (entity) {
		const key = entity.toLowerCase();
		return candidates.find((a) => a.entity?.alias?.toLowerCase() === key || a.entity?.name?.toLowerCase() === key);
	}
	return candidates.find((a) => a.entity?.isRoot) ?? candidates[0];
}

/** The node a Dataverse-only finding points at: the first condition / order on the attribute, else the root. */
function locateIssue(fetch: FetchNode, rule: AdvisorRuleId | undefined, attribute: string | undefined): NodeId {
	if (!attribute || !rule) return fetch.entity.id;
	const want = attribute.toLowerCase();
	const containers: Container[] = [];
	const conditions: ConditionNode[] = [];
	const visit = (c: Container) => {
		containers.push(c);
		const visitFilter = (f: FilterNode) => {
			conditions.push(...f.conditions);
			f.subfilters.forEach(visitFilter);
			f.links?.forEach(visit);
		};
		c.filters.forEach(visitFilter);
		c.links.forEach(visit);
	};
	visit(fetch.entity);
	if (rule === "order-choice" || rule === "order-link") {
		const order = containers.flatMap((c) => c.orders).find((o) => o.attribute?.toLowerCase() === want);
		return order?.id ?? fetch.entity.id;
	}
	// A condition the user is still editing can have no attribute yet (#217): never throw here
	return conditions.find((c) => c.attribute?.toLowerCase() === want)?.id ?? fetch.entity.id;
}

export interface DataverseCheck {
	issues: readonly DataverseIssue[];
	/** ValidationResults.Helplink */
	helplink?: string;
}

/**
 * Merge Dataverse's findings into FetchXML Studio's: a finding Dataverse also reported (same rule
 * and attribute) becomes "Dataverse agrees"; anything else becomes a "From Dataverse" row quoting
 * Dataverse's message. ValidateFetchXmlExpression reports performance issues, not rejections, so
 * its rows are advice. Returns the merged list and how many rows only Dataverse reported.
 */
export function mergeDataverseCheck(
	fetch: FetchNode,
	advisories: readonly Advisory[],
	check: DataverseCheck,
): { advisories: Advisory[]; added: number } {
	const merged = advisories.map((a) => ({ ...a }));
	const locations = nodeLocations(fetch);
	const added: Advisory[] = [];
	check.issues.forEach((issue, index) => {
		const text = issue.LocalizedMessageText ?? "";
		const rules = DATAVERSE_TYPE_RULES[issue.TypeCode];
		const attribute = issueAttribute(issue);
		const queryLevel = QUERY_LEVEL_TYPES.has(issue.TypeCode);
		if (rules && (attribute || queryLevel)) {
			const match = findAgreeing(merged, rules, attribute, issueEntity(issue), queryLevel);
			if (match) {
				match.source = "both";
				match.dataverseText = text;
				return;
			}
			const rule = rules[0];
			const nodeId = locateIssue(fetch, rule, attribute);
			added.push({
				id: `dataverse:${index}`,
				level: "advice",
				title: RULES[rule].title,
				detail: attribute ? `Dataverse reported this for '${attribute}'.` : "Dataverse reported this for the query.",
				nodeId,
				location: locations.get(nodeId) ?? fetch.entity.name,
				doc: RULES[rule].doc,
				source: "dataverse",
				dataverseText: text,
				rule,
				target: attribute,
			});
			return;
		}
		// An unknown code, or a message without the attribute: shown as Dataverse wrote it
		added.push({
			id: `dataverse:${index}`,
			level: "advice",
			title: "Reported by Dataverse",
			detail: "Dataverse's query check reported this.",
			nodeId: fetch.entity.id,
			location: locations.get(fetch.entity.id) ?? fetch.entity.name,
			doc: check.helplink ? { label: "Dataverse help", href: check.helplink } : { label: ANTIPATTERNS_LABEL, href: ANTIPATTERNS },
			source: "dataverse",
			dataverseText: text,
		});
	});
	return { advisories: [...merged, ...added].sort((a, b) => LEVEL_ORDER[a.level] - LEVEL_ORDER[b.level]), added: added.length };
}

/** The entities whose metadata the metadata rules use: the root, then every link-entity (also inside filters). */
export function advisorEntityNames(fetch: FetchNode): string[] {
	const names = new Set<string>();
	const visit = (c: Container) => {
		if (c.name) names.add(c.name);
		const visitFilter = (f: FilterNode) => {
			f.subfilters.forEach(visitFilter);
			f.links?.forEach(visit);
		};
		c.filters.forEach(visitFilter);
		c.links.forEach(visit);
	};
	visit(fetch.entity);
	return [...names];
}

/** How many advisories there are per level */
export function countByLevel(list: readonly Advisory[]): Record<AdvisoryLevel, number> {
	const counts: Record<AdvisoryLevel, number> = { error: 0, warning: 0, advice: 0 };
	list.forEach((a) => counts[a.level]++);
	return counts;
}
