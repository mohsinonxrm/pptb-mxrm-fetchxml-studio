/**
 * Query Advisor (#97): documented anti-patterns and limits. Each rule has positive and negative
 * cases; advice never blocks Execute (no errors, no "Execute anyway" warnings).
 */
import { describe, expect, it } from "vitest";
import { getBlockingWarnings, getExecutionErrors, parseFetchXml } from "./fetchxmlParser";
import { adviseCondition, adviseQuery, hasLeadingWildcard } from "./queryAdvisor";
import type { RuleFinding } from "./fetchxmlRules";
import type { ConditionNode } from "./nodes";

const ACC = (inner: string, fetchAttrs = "") =>
	`<fetch${fetchAttrs}><entity name="account"><attribute name="name" />${inner}</entity></fetch>`;
const COND = (operator: string, value?: string) =>
	ACC(`<filter><condition attribute="name" operator="${operator}"${value === undefined ? "" : ` value="${value}"`} /></filter>`);
const CONTACT = (inner = "", alias = ` alias="c"`) =>
	`<link-entity name="contact" from="parentcustomerid" to="accountid"${alias}>${inner}</link-entity>`;
const AGG = (inner = "", fetchAttrs = "") =>
	`<fetch aggregate="true"${fetchAttrs}><entity name="account"><attribute name="accountid" alias="n" aggregate="count" />${inner}</entity></fetch>`;

const advice = (xml: string): RuleFinding[] => {
	const r = parseFetchXml(xml);
	expect(r.success, r.errors.map((e) => e.message).join("; ")).toBe(true);
	return adviseQuery(r.fetchNode!);
};
/** The advice for one rule, identified by element + attribute */
const ruleOf = (xml: string, element: string, attribute?: string) =>
	advice(xml).filter((f) => f.element === element && f.attribute === attribute);

type Case = [label: string, xml: string, element: string, attribute: string | undefined, needle: string];

const ADVICE_CASES: Case[] = [
	["like with a leading %", COND("like", "%234"), "condition", "value", "0x80048644"],
	["like with a leading _", COND("like", "_234%"), "condition", "value", "starts with a wildcard"],
	["like with a leading [^a]", COND("like", "[^a]234%"), "condition", "value", "starts with a wildcard"],
	["not-like with a leading %", COND("not-like", "%abc"), "condition", "value", "starts with a wildcard"],
	["begins-with with a leading %", COND("begins-with", "%value"), "condition", "value", "starts with a wildcard"],
	["not-begin-with with a leading %", COND("not-begin-with", "%value"), "condition", "value", "starts with a wildcard"],
	["like '-%234' (hyphen before a wildcard)", COND("like", "-%234"), "condition", "value", "hyphen or apostrophe"],
	["like with an apostrophe before a wildcard", COND("like", "&apos;%751"), "condition", "value", "hyphen or apostrophe"],
	["begins-with '-' (implicit trailing %)", COND("begins-with", "-"), "condition", "value", "hyphen or apostrophe"],
	["ends-with", COND("ends-with", "value"), "condition", "operator", "works like a leading wildcard"],
	["not-end-with", COND("not-end-with", "value"), "condition", "operator", "works like a leading wildcard"],
	["all-attributes on the root entity", ACC("<all-attributes />"), "all-attributes", undefined, "Select only the attributes you need"],
	["all-attributes on a link-entity", ACC(CONTACT("<all-attributes />")), "all-attributes", undefined, "All attributes of contact"],
	["no attributes on the root entity", `<fetch><entity name="account"><filter><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>`, "entity", undefined, "returns all of them"],
	["order inside a link-entity", ACC(CONTACT(`<attribute name="fullname" /><order attribute="fullname" />`)), "order", "attribute", "paging cookie"],
	["root order with entityname = link alias", ACC(`${CONTACT(`<attribute name="fullname" />`)}<order entityname="c" attribute="fullname" />`), "order", "entityname", "Sort by the root entity's attributes"],
	["root order with entityname = link name", ACC(`${CONTACT(`<attribute name="fullname" />`, "")}<order entityname="contact" attribute="fullname" />`), "order", "entityname", "slower"],
	["distinct without an order", ACC("", ` distinct="true"`), "fetch", "distinct", "add at least one order"],
	["simple paging past 50,000 records", ACC(`<order attribute="name" />`, ` count="5000" page="11"`), "fetch", "page", "50,000"],
	["simple paging past 50,000 records with the default page size", ACC(`<order attribute="name" />`, ` page="11"`), "fetch", "page", "5000 records per page"],
	["aggregate query without a filter", AGG(), "fetch", "aggregate", "0x8004E023"],
	["no-lock", ACC("", ` no-lock="true"`), "fetch", "no-lock", "no longer necessary"],
];

const NO_ADVICE_CASES: Case[] = [
	["like with a trailing %", COND("like", "abc%"), "condition", "value", ""],
	["like with a wildcard after literal text", COND("like", "a%b_c"), "condition", "value", ""],
	["like '-234%' (literal text before the wildcard)", COND("like", "-234%"), "condition", "value", ""],
	["begins-with literal text", COND("begins-with", "Contoso"), "condition", "value", ""],
	["begins-with '-234'", COND("begins-with", "-234"), "condition", "value", ""],
	["eq with a leading % (not a pattern)", COND("eq", "%234"), "condition", "value", ""],
	["like compared to another attribute (valueof)", ACC(`<filter><condition attribute="name" operator="like" valueof="accountnumber" /></filter>`), "condition", "value", ""],
	["begins-with has no operator advice", COND("begins-with", "abc"), "condition", "operator", ""],
	["selected attributes only", ACC(""), "all-attributes", undefined, ""],
	["all-attributes in an aggregate query", AGG("<all-attributes />"), "all-attributes", undefined, ""],
	["root entity with attributes", ACC(""), "entity", undefined, ""],
	["link-entity without attributes (returns none)", ACC(CONTACT()), "entity", undefined, ""],
	["root order", ACC(`<order attribute="name" />`), "order", "attribute", ""],
	["root order with entityname = root entity", ACC(`<order entityname="account" attribute="name" />`), "order", "entityname", ""],
	["aggregate order inside a link-entity", AGG(CONTACT(`<attribute name="fullname" alias="f" groupby="true" /><order alias="f" />`)), "order", "alias", ""],
	["distinct with an order", ACC(`<order attribute="name" />`, ` distinct="true"`), "fetch", "distinct", ""],
	["distinct in an aggregate query", AGG("", ` distinct="true"`), "fetch", "distinct", ""],
	["simple paging up to 50,000 records", ACC(`<order attribute="name" />`, ` count="5000" page="10"`), "fetch", "page", ""],
	["paging with a paging cookie", ACC(`<order attribute="name" />`, ` count="5000" page="11" paging-cookie="&lt;cookie page=&quot;10&quot; /&gt;"`), "fetch", "page", ""],
	["aggregate query with a filter", AGG(`<filter><condition attribute="statecode" operator="eq" value="0" /></filter>`), "fetch", "aggregate", ""],
	["aggregate query with an aggregate row limit", AGG("", ` aggregatelimit="1000"`), "fetch", "aggregate", ""],
	["aggregate query filtered by a link-entity in a filter", AGG(`<filter><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="any" /></filter>`), "fetch", "aggregate", ""],
	["non-aggregate query without a filter", ACC(""), "fetch", "aggregate", ""],
	["no no-lock", ACC(""), "fetch", "no-lock", ""],
];

describe("Query Advisor: documented anti-patterns → advice", () => {
	it.each(ADVICE_CASES)("%s", (_label, xml, element, attribute, needle) => {
		const found = ruleOf(xml, element, attribute);
		expect(found).toHaveLength(1);
		expect(found[0]).toMatchObject({ severity: "info", kind: "advice" });
		expect(found[0].message).toContain(needle);
	});

	it.each(NO_ADVICE_CASES)("no advice: %s", (_label, xml, element, attribute) => {
		expect(ruleOf(xml, element, attribute)).toEqual([]);
	});
});

describe("Query Advisor: advice never blocks Execute", () => {
	it.each(ADVICE_CASES)("%s", (_label, xml) => {
		const result = parseFetchXml(xml);
		const advisor = result.warnings.filter((w) => w.kind === "advice" && w.nodeId !== undefined);
		expect(advisor.length).toBeGreaterThan(0);
		expect(advisor.every((w) => w.severity === "info")).toBe(true);
		expect(getExecutionErrors(result)).toEqual([]);
		expect(getBlockingWarnings(result)).toEqual([]);
	});
});

describe("hasLeadingWildcard", () => {
	it.each([
		["like", "%a", true],
		["like", "_a", true],
		["like", "[a-c]x", true],
		["like", "-%", true],
		["like", "'%", true],
		["like", "--_x", true],
		["begins-with", "'", true],
		["like", "a%", false],
		["like", "-a%", false],
		["like", "-", false],
		["begins-with", "a", false],
		["eq", "%a", false],
		["ends-with", "%a", false], // reported as an operator, not a value
		["like", "", false],
		// Bracket escapes are literal characters (#98, wildcard-characters)
		["like", "[%]a%", false],
		["like", "[_]a", false],
		["like", "[[]a]%", false],
		["begins-with", "[%]", false],
		["like", "-[%]%", false],
		["like", "%[%]%", true],
		["like", "[%a]", true],
	] as const)("%s '%s' → %s", (operator, value, expected) => {
		expect(hasLeadingWildcard(operator, value)).toBe(expected);
	});

	it("ignores non-string values", () => {
		expect(hasLeadingWildcard("like", 5)).toBe(false);
		expect(hasLeadingWildcard("like", ["%a"])).toBe(false);
	});
});

describe("adviseCondition (column filters, #98)", () => {
	const cond = (operator: string, value?: string, valueof?: string): ConditionNode => ({
		id: "c1", type: "condition", attribute: "name", operator: operator as ConditionNode["operator"], value, valueof,
	});
	it("gives the same advice as the query for one condition", () => {
		expect(adviseCondition(cond("like", "%abc%"))).toMatchObject([{ attribute: "value", severity: "info", kind: "advice" }]);
		expect(adviseCondition(cond("ends-with", "abc"))).toMatchObject([{ attribute: "operator" }]);
		expect(adviseCondition(cond("like", "abc%"))).toEqual([]);
		expect(adviseCondition(cond("like", "[%]abc%"))).toEqual([]);
		expect(adviseCondition(cond("ends-with", undefined, "other"))).toEqual([]);
	});
});

const links = (n: number, linkType = "inner") =>
	Array.from({ length: n }, (_, i) => `<link-entity name="contact" from="parentcustomerid" to="accountid" alias="l${i}" link-type="${linkType}" />`).join("");
const conditions = (n: number) =>
	`<filter>${Array.from({ length: n }, (_, i) => `<condition attribute="name" operator="eq" value="v${i}" />`).join("")}</filter>`;

describe("Query Advisor: documented limits", () => {
	it("15 joined link-entities: no finding", () => {
		expect(advice(ACC(links(15))).filter((f) => f.message.includes("0x8004430D"))).toEqual([]);
	});

	it("16 joined link-entities: error on the 16th (0x8004430D)", () => {
		const r = parseFetchXml(ACC(links(16)));
		const found = adviseQuery(r.fetchNode!).filter((f) => f.message.includes("0x8004430D"));
		expect(found).toHaveLength(1);
		expect(found[0]).toMatchObject({ element: "link-entity", severity: "error", kind: "invalid", nodeId: r.fetchNode!.entity.links[15].id });
		expect(getExecutionErrors(r).some((w) => w.message.includes("0x8004430D"))).toBe(true);
	});

	it("nested joined link-entities count toward the limit", () => {
		let nested = "";
		for (let i = 15; i >= 0; i--) nested = `<link-entity name="contact" from="parentcustomerid" to="accountid" alias="n${i}">${nested}</link-entity>`;
		expect(advice(ACC(nested)).filter((f) => f.severity === "error" && f.message.includes("0x8004430D"))).toHaveLength(1);
	});

	it("link-entities inside filters don't count toward 15 (verified live: 14 joined + 2 any run)", () => {
		const xml = ACC(`${links(14)}<filter>${links(2, "any").replace(/alias="l/g, 'alias="f')}</filter>`);
		expect(advice(xml).filter((f) => f.message.includes("0x8004430D"))).toEqual([]);
		expect(getExecutionErrors(parseFetchXml(xml))).toEqual([]);
	});

	it("500 conditions plus a link-entity: no finding (link-entities don't count, verified live)", () => {
		expect(advice(ACC(`${conditions(500)}${links(1)}`)).filter((f) => f.message.includes("0x8004430C"))).toEqual([]);
	});

	it("501 conditions: error (0x8004430C, verified live)", () => {
		const found = advice(ACC(conditions(501))).filter((f) => f.message.includes("0x8004430C"));
		expect(found).toHaveLength(1);
		expect(found[0]).toMatchObject({ element: "fetch", severity: "error", kind: "invalid" });
		expect(found[0].message).toContain("501");
	});

	it("URL-encoded FetchXML over 32,768 characters: warning (0x80060888)", () => {
		const found = advice(COND("eq", "x".repeat(33_000))).filter((f) => f.message.includes("0x80060888"));
		expect(found).toHaveLength(1);
		expect(found[0]).toMatchObject({ element: "fetch", severity: "warning", kind: "invalid" });
	});

	it("short FetchXML: no URL length finding", () => {
		expect(advice(COND("eq", "x".repeat(1000))).filter((f) => f.message.includes("0x80060888"))).toEqual([]);
	});
});
