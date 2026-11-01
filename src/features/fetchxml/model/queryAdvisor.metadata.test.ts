/**
 * Query Advisor metadata rules (#97): choice orders, large text and calculated attributes in
 * conditions, and paging without a unique order. Each rule has positive and negative cases, and is
 * skipped when the metadata it needs isn't loaded.
 */
import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { adviseQueryMetadata, largeTextReason, type AdvisorAttributeInfo, type AdvisorMetadata } from "./queryAdvisor";
import type { FetchNode } from "./nodes";

const attrs = (entries: Record<string, AdvisorAttributeInfo>) => new Map(Object.entries(entries));

const CONTACT_ATTRIBUTES = attrs({
	contactid: { typeName: "UniqueidentifierType" },
	fullname: { typeName: "StringType", maxLength: 160 },
	statuscode: { typeName: "StatusType" },
	statecode: { typeName: "StateType" },
	preferredcontactmethodcode: { typeName: "PicklistType" },
	donotemail: { typeName: "BooleanType" },
	parentcustomerid: { typeName: "CustomerType" },
	description: { typeName: "MemoType", maxLength: 2000 },
	traversedpath: { typeName: "StringType", maxLength: 1250 },
	exactly850: { typeName: "StringType", maxLength: 850 },
	nolength: { typeName: "StringType" },
	calc: { typeName: "StringType", maxLength: 100, sourceType: 1 },
	rollup: { typeName: "IntegerType", sourceType: 2 },
	formula: { typeName: "DecimalType", sourceType: 3 },
	simple: { typeName: "IntegerType", sourceType: 0 },
});
const ACCOUNT_ATTRIBUTES = attrs({
	accountid: { typeName: "UniqueidentifierType" },
	name: { typeName: "StringType", maxLength: 160 },
	industrycode: { typeName: "PicklistType" },
	description: { typeName: "MemoType" },
});

const METADATA: AdvisorMetadata = new Map([
	["contact", { primaryIdAttribute: "contactid", attributes: CONTACT_ATTRIBUTES }],
	["account", { primaryIdAttribute: "accountid", attributes: ACCOUNT_ATTRIBUTES }],
]);

const parse = (xml: string): FetchNode => {
	const r = parseFetchXml(xml);
	expect(r.success, r.errors.map((e) => e.message).join("; ")).toBe(true);
	return r.fetchNode!;
};
const advise = (xml: string, metadata: AdvisorMetadata = METADATA) => adviseQueryMetadata(parse(xml), metadata);
const rule = (xml: string, id: string, metadata?: AdvisorMetadata) => advise(xml, metadata).filter((f) => f.rule === id);

const C = (inner: string, fetchAttrs = "") =>
	`<fetch${fetchAttrs}><entity name="contact"><attribute name="fullname" />${inner}</entity></fetch>`;
const LINK = (inner = "", alias = ` alias="a"`) =>
	`<link-entity name="account" from="accountid" to="parentcustomerid"${alias}><attribute name="name" />${inner}</link-entity>`;
const COND = (attribute: string, operator = "eq", value: string | undefined = "x") =>
	`<filter><condition attribute="${attribute}" operator="${operator}"${value === undefined ? "" : ` value="${value}"`} /></filter>`;

describe("order on a choice attribute", () => {
	it.each([
		["Status", "statuscode", "choice"],
		["State", "statecode", "choice"],
		["Picklist", "preferredcontactmethodcode", "choice"],
		["Boolean (Yes/No)", "donotemail", "Yes/No"],
	])("%s → advice", (_label, attribute, kind) => {
		const found = rule(C(`<order attribute="${attribute}" />`), "order-choice");
		expect(found).toHaveLength(1);
		expect(found[0]).toMatchObject({ element: "order", severity: "info", kind: "advice", target: attribute });
		expect(found[0].message).toContain(`${kind} attribute '${attribute}'`);
		expect(found[0].message).toContain("useraworderby");
	});

	it.each([
		["a String attribute", C(`<order attribute="fullname" />`)],
		["a lookup", C(`<order attribute="parentcustomerid" />`)],
		["useraworderby (sorts by the number; verified live: no TypeCode 4)", C(`<order attribute="statuscode" />`, ` useraworderby="true"`)],
		["an aggregate query (orders by alias)", `<fetch aggregate="true"><entity name="contact"><attribute name="statuscode" alias="s" groupby="true" /><order alias="s" /></entity></fetch>`],
		["an attribute missing from the metadata", C(`<order attribute="unknown" />`)],
	])("no advice: %s", (_label, xml) => {
		expect(rule(xml, "order-choice")).toEqual([]);
	});

	it("uses the link-entity's metadata for an order inside it, and for entityname orders", () => {
		expect(rule(C(LINK(`<order attribute="industrycode" />`)), "order-choice")).toHaveLength(1);
		expect(rule(C(`${LINK()}<order entityname="a" attribute="industrycode" />`), "order-choice")).toHaveLength(1);
		expect(rule(C(`${LINK("", "")}<order entityname="account" attribute="industrycode" />`), "order-choice")).toHaveLength(1);
		expect(rule(C(LINK(`<order attribute="name" />`)), "order-choice")).toEqual([]);
	});

	it("is skipped when the entity's metadata isn't loaded", () => {
		expect(rule(C(`<order attribute="statuscode" />`), "order-choice", new Map())).toEqual([]);
		expect(rule(C(`<order attribute="statuscode" />`), "order-choice", new Map([["contact", { primaryIdAttribute: "contactid" }]]))).toEqual([]);
	});
});

describe("condition on a large text attribute", () => {
	it.each([
		["Memo, like", COND("description", "like", "%vip%"), "Memo"],
		["Memo, eq (any operator, verified live)", COND("description", "eq", "x"), "Memo"],
		["Memo, not-null", COND("description", "not-null", undefined), "Memo"],
		["Memo, begins-with", COND("description", "begins-with", "x"), "Memo"],
		["String with MaxLength 1,250", COND("traversedpath", "like", "x%"), "maximum length of 1,250"],
	])("%s → advice", (_label, filter, needle) => {
		const found = rule(C(filter), "large-text");
		expect(found).toHaveLength(1);
		expect(found[0]).toMatchObject({ element: "condition", severity: "info", kind: "advice" });
		expect(found[0].message).toContain(needle);
		expect(found[0].message).toContain("Dataverse search");
	});

	it.each([
		["String with MaxLength 160", COND("fullname", "like", "x%")],
		["String with MaxLength exactly 850 (not greater than 850)", COND("exactly850")],
		["String whose MaxLength isn't loaded (never guessed)", COND("nolength")],
		["an attribute missing from the metadata", COND("unknown")],
	])("no advice: %s", (_label, filter) => {
		expect(rule(C(filter), "large-text")).toEqual([]);
	});

	it("uses the link-entity's metadata, for its own filters and for entityname conditions", () => {
		expect(rule(C(LINK(COND("description", "like", "x%"))), "large-text")).toHaveLength(1);
		expect(rule(C(`${LINK()}<filter><condition entityname="a" attribute="description" operator="eq" value="x" /></filter>`), "large-text")).toHaveLength(1);
		expect(rule(C(`<filter><link-entity name="account" from="accountid" to="parentcustomerid" link-type="any">${COND("description")}</link-entity></filter>`), "large-text")).toHaveLength(1);
	});

	it("largeTextReason", () => {
		expect(largeTextReason({ typeName: "MemoType" })).toContain("Memo");
		expect(largeTextReason({ typeName: "StringType", maxLength: 851 })).toContain("851");
		expect(largeTextReason({ typeName: "StringType", maxLength: 850 })).toBeUndefined();
		expect(largeTextReason({ typeName: "StringType" })).toBeUndefined();
		expect(largeTextReason(undefined)).toBeUndefined();
	});
});

describe("condition on a calculated or formula attribute", () => {
	it.each([
		["calculated (SourceType 1)", "calc", "calculated"],
		["formula (SourceType 3)", "formula", "formula"],
	])("%s → advice", (_label, attribute, kind) => {
		const found = rule(C(COND(attribute, "eq", "1")), "calculated-filter");
		expect(found).toHaveLength(1);
		expect(found[0].message).toContain(`is a ${kind} attribute`);
		expect(found[0].message).toContain("0x80048744");
	});

	it.each([
		["rollup (SourceType 2, stored; verified live: no TypeCode 8)", "rollup"],
		["simple (SourceType 0)", "simple"],
		["no SourceType", "fullname"],
	])("no advice: %s", (_label, attribute) => {
		expect(rule(C(COND(attribute, "eq", "1")), "calculated-filter")).toEqual([]);
	});
});

describe("paging without a unique order", () => {
	it("count + page with orders but none on the primary key → advice with the quick fix", () => {
		const found = rule(C(`<order attribute="fullname" />`, ` count="50" page="1"`), "paging-order");
		expect(found).toHaveLength(1);
		expect(found[0]).toMatchObject({ element: "entity", severity: "info", kind: "advice", target: "contactid", fix: { kind: "add-root-order", attribute: "contactid" } });
		expect(found[0].message).toContain("Add contactid as the last order");
	});

	it("count alone is paging too", () => {
		expect(rule(C(`<order attribute="fullname" />`, ` count="50"`), "paging-order")).toHaveLength(1);
	});

	it("orders inside a link-entity: advice without the quick fix (link-entity orders are applied last)", () => {
		const found = rule(C(`<order attribute="statuscode" />${LINK(`<order attribute="name" />`)}`, ` count="50" page="1"`), "paging-order");
		expect(found).toHaveLength(1);
		expect(found[0].fix).toBeUndefined();
		expect(found[0].message).toContain("entityname");
	});

	it.each([
		["the primary key is ordered", C(`<order attribute="fullname" /><order attribute="contactid" />`, ` count="50" page="1"`)],
		["the primary key is ordered with entityname = the root", C(`<order attribute="fullname" /><order entityname="contact" attribute="contactid" />`, ` count="50"`)],
		["no order (Dataverse adds the primary key, verified live)", C("", ` count="50" page="1"`)],
		["not paged", C(`<order attribute="fullname" />`)],
		["top", C(`<order attribute="fullname" />`, ` top="50"`)],
		["distinct (its own rule)", C(`<order attribute="fullname" />`, ` distinct="true" count="50"`)],
		["aggregate", `<fetch aggregate="true" count="50"><entity name="contact"><attribute name="statuscode" alias="s" groupby="true" /><order alias="s" /></entity></fetch>`],
	])("no advice: %s", (_label, xml) => {
		expect(rule(xml, "paging-order")).toEqual([]);
	});

	it("is skipped when the primary key isn't known", () => {
		const metadata: AdvisorMetadata = new Map([["contact", { attributes: CONTACT_ATTRIBUTES }]]);
		expect(rule(C(`<order attribute="fullname" />`, ` count="50"`), "paging-order", metadata)).toEqual([]);
	});

	it("an order on the primary key inside a link-entity doesn't count", () => {
		expect(rule(C(LINK(`<order attribute="accountid" />`), ` count="50"`), "paging-order")).toHaveLength(1);
	});

	it("paged results count as paging (Load more, Retrieve all, #170)", () => {
		const xml = C(`<order attribute="fullname" />`);
		expect(adviseQueryMetadata(parse(xml), METADATA).filter((f) => f.rule === "paging-order")).toEqual([]);
		const found = adviseQueryMetadata(parse(xml), METADATA, { resultsPaged: true }).filter((f) => f.rule === "paging-order");
		expect(found).toHaveLength(1);
		expect(found[0].message).toMatch(/^The results are paged, but no order is on a unique attribute such as contactid/);
		expect(found[0].fix).toEqual({ kind: "add-root-order", attribute: "contactid" });
		// top caps the results, so the query doesn't page however it was run
		expect(adviseQueryMetadata(parse(C(`<order attribute="fullname" />`, ` top="50"`)), METADATA, { resultsPaged: true }).filter((f) => f.rule === "paging-order")).toEqual([]);
	});
});

describe("the prototype's scenario query", () => {
	const SCENARIO = `<fetch count="50" page="1"><entity name="contact"><attribute name="fullname" /><attribute name="statuscode" /><order attribute="statuscode" /><filter><condition attribute="fullname" operator="like" value="%smith%" /><condition attribute="description" operator="like" value="%vip%" /></filter><link-entity name="account" from="accountid" to="parentcustomerid" alias="a"><attribute name="name" /><order attribute="name" /></link-entity></entity></fetch>`;

	it("finds the choice order, the large text condition and the paging order; all advice", () => {
		const found = advise(SCENARIO);
		expect(found.map((f) => `${f.rule}:${f.target}`).sort()).toEqual(["large-text:description", "order-choice:statuscode", "paging-order:contactid"]);
		expect(found.every((f) => f.severity === "info" && f.kind === "advice")).toBe(true);
	});
});
