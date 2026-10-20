/**
 * Incomplete nodes never crash (1.4.3 hotfix). The property editors write undefined into fields the
 * node types call required while the user is mid-edit: picking a condition's attribute clears its
 * operator, changing its entity clears its attribute, and clearing a picker clears the field. 1.4.2
 * crashed the whole tool on the first of these ("Cannot read properties of undefined (reading
 * 'replace')" in escapeXmlAttr), because the FetchXML preview serializes the tree on every render.
 *
 * Every pure consumer of the tree must cope with any string field being undefined or empty. This
 * test blanks each string field of every node in a rich query, one at a time, and runs them all.
 */
import { describe, expect, it } from "vitest";
import { generateFetchXml } from "./fetchxml";
import { parseFetchXml } from "./fetchxmlParser";
import { conditionValueIssue, validateQueryRules, valueofScopes } from "./fetchxmlRules";
import { planAggregateOff } from "./aggregateToggle";
import { collectLinkEntityAliases } from "./aliasUtils";
import { columnFilterTarget, columnFilterConditions, findColumnFilterGroup } from "./columnFilters";
import {
	collectQueryColumns,
	joinedLinkAliases,
	resolveColumn,
	sortedColumns,
} from "./queryColumns";
import {
	collectAllFilters,
	collectAllLinkEntities,
	collectLinkEntityReferences,
	getQueryEntities,
} from "./treeUtils";
import { generateFxsSyncCode, type FxsSyncFormat } from "../engine/fetchxmlCodeGenerators";
import type { FetchNode } from "./nodes";
import {
	advisorEntityNames,
	collectAdvisories,
	mergeDataverseCheck,
	nodeEntities,
	nodeLocations,
} from "./advisories";
import { adviseQueryMetadata, type AdvisorMetadata } from "./queryAdvisor";

/** Metadata for the Advisories metadata rules (1.6, #97): choice, Memo, calculated and the primary key */
const ADVISOR_METADATA: AdvisorMetadata = new Map([
	[
		"account",
		{
			primaryIdAttribute: "accountid",
			attributes: new Map([
				["name", { typeName: "StringType", maxLength: 160 }],
				["statecode", { typeName: "StateType" }],
				["revenue", { typeName: "MoneyType", sourceType: 1 }],
				["description", { typeName: "MemoType" }],
			]),
		},
	],
	[
		"contact",
		{
			primaryIdAttribute: "contactid",
			attributes: new Map([["fullname", { typeName: "StringType", maxLength: 160 }]]),
		},
	],
]);

const RICH = `<fetch top="50" distinct="true" no-lock="true" options="OptimizeForUnknown">
  <entity name="account">
    <attribute name="name" alias="n" />
    <attribute name="accountnumber" />
    <order attribute="name" descending="true" />
    <filter type="and" hint="union">
      <condition attribute="name" operator="like" value="%contoso%" />
      <condition attribute="statecode" operator="in"><value>0</value><value>1</value></condition>
      <condition attribute="revenue" operator="between"><value>1</value><value>100</value></condition>
      <condition attribute="numberofemployees" operator="gt" valueof="revenue" />
      <condition entityname="c" attribute="fullname" operator="not-null" />
      <filter type="or">
        <condition attribute="createdon" operator="last-x-days" value="7" />
        <link-entity name="task" from="regardingobjectid" to="accountid" link-type="exists" alias="t">
          <filter><condition attribute="subject" operator="eq" value="x" /></filter>
        </link-entity>
      </filter>
    </filter>
    <link-entity name="contact" from="parentcustomerid" to="accountid" link-type="outer" alias="c">
      <attribute name="fullname" />
      <order attribute="fullname" />
      <filter><condition attribute="emailaddress1" operator="ends-with" value="@contoso.com" /></filter>
    </link-entity>
  </entity>
</fetch>`;

const FORMATS: FxsSyncFormat[] = [
	"csharp-fetchexpr",
	"javascript",
	"pac",
	"powerautomate",
	"webapi",
];

/** Every object in the tree, with a path for messages */
function nodesOf(
	value: unknown,
	path = "fetch",
	out: { path: string; node: Record<string, unknown> }[] = [],
) {
	if (Array.isArray(value)) value.forEach((v, i) => nodesOf(v, `${path}[${i}]`, out));
	else if (value && typeof value === "object") {
		const node = value as Record<string, unknown>;
		if (typeof node.type === "string") out.push({ path, node });
		for (const [k, v] of Object.entries(node))
			if (v && typeof v === "object") nodesOf(v, `${path}.${k}`, out);
	}
	return out;
}

/** Run every pure consumer; throws if any of them does */
function consumeAll(fetch: FetchNode) {
	const xml = generateFetchXml(fetch);
	validateQueryRules(fetch);
	valueofScopes(fetch);
	collectAllFilters(fetch).forEach((f) => f.conditions.forEach((c) => conditionValueIssue(c)));
	planAggregateOff(fetch);
	planAggregateOff({ ...fetch, options: { ...fetch.options, aggregate: true } });
	collectLinkEntityAliases(fetch);
	collectQueryColumns(fetch).forEach((c) => {
		resolveColumn(fetch, c.key);
		columnFilterTarget(fetch, c.key);
	});
	joinedLinkAliases(fetch);
	sortedColumns(fetch);
	collectLinkEntityReferences(fetch);
	getQueryEntities(fetch);
	collectAllLinkEntities(fetch);
	collectAllFilters(fetch).forEach((f) => {
		findColumnFilterGroup(fetch, f.id);
		columnFilterConditions(fetch, f.id);
	});
	// The Advisories tab (1.6, #97): its list, metadata rules and the Dataverse merge
	nodeLocations(fetch);
	nodeEntities(fetch);
	advisorEntityNames(fetch);
	adviseQueryMetadata(fetch, ADVISOR_METADATA);
	const advisories = collectAdvisories(fetch, ADVISOR_METADATA);
	mergeDataverseCheck(fetch, advisories, {
		issues: [
			{
				TypeCode: 2,
				Severity: 1,
				LocalizedMessageText:
					"Leading wild card on like/notlike conditions can lead to slow performance. Attribute name",
			},
			{
				TypeCode: 4,
				Severity: 1,
				LocalizedMessageText:
					"Ordering on Enum variables can lead to slow performance. Attribute statecode",
			},
			{ TypeCode: 99, Severity: 1, LocalizedMessageText: "Something new" },
		],
	});
	// The editor round trip: what the preview shows must parse again
	parseFetchXml(xml);
	FORMATS.forEach((fmt) =>
		generateFxsSyncCode(fmt, xml, "accounts", "https://contoso.crm.dynamics.com"),
	);
	return xml;
}

describe("incomplete nodes never crash the tool (1.4.3)", () => {
	const parsed = parseFetchXml(RICH);
	it("parses the rich query", () => {
		expect(parsed.success).toBe(true);
		expect(() => consumeAll(parsed.fetchNode!)).not.toThrow();
	});

	const base = parsed.fetchNode!;
	const cases: { label: string; mutate: (fetch: FetchNode) => void }[] = [];
	nodesOf(base).forEach(({ path, node }, index) => {
		for (const [key, value] of Object.entries(node)) {
			if (typeof value !== "string" || key === "id" || key === "type") continue;
			for (const blank of [undefined, ""] as const) {
				cases.push({
					label: `${path} (${node.type}).${key} = ${blank === undefined ? "undefined" : '""'}`,
					mutate: (fetch) => {
						const target = nodesOf(fetch)[index].node;
						target[key] = blank;
					},
				});
			}
		}
	});

	it("has a case for every string field", () => {
		expect(cases.length).toBeGreaterThan(60);
	});

	it.each(cases)("$label", ({ mutate }) => {
		const fetch = structuredClone(base);
		mutate(fetch);
		expect(() => consumeAll(fetch)).not.toThrow();
	});

	// The exact editor updates that crashed 1.4.2
	it.each([
		[
			"pick an attribute (operator cleared)",
			{ attribute: "telephone1", operator: undefined, value: undefined },
		],
		[
			"change the condition's entity (attribute and operator cleared)",
			{ entityname: "c", attribute: undefined, operator: undefined, value: undefined },
		],
		[
			"clear the attribute picker",
			{ attribute: undefined, operator: undefined, value: undefined },
		],
	])("ConditionEditor: %s", (_label, update) => {
		const fetch = structuredClone(base);
		const condition = fetch.entity.filters[0].conditions[0];
		Object.assign(condition, update);
		const xml = consumeAll(fetch);
		expect(xml).toContain("<condition");
		expect(xml).not.toContain("undefined");
	});
});
