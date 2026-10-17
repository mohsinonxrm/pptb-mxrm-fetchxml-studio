/**
 * Query-level rules (#90b). Every case mirrors a query executed against a live Dataverse org
 * through the PPTB request path (2026-10-13/14); the expected outcome is what Dataverse did.
 */
import { describe, expect, it } from "vitest";
import { parseFetchXml, getExecutionErrors, type ParseWarning } from "./fetchxmlParser";
import { validateQueryRules } from "./fetchxmlRules";

const SU = (inner: string, fetchAttrs = "") => `<fetch${fetchAttrs}><entity name="systemuser">${inner}</entity></fetch>`;
const CNT = `<attribute name="systemuserid" alias="n" aggregate="count" />`;
const GRP = `<attribute name="isdisabled" alias="d" groupby="true" />`;
const BU = (inner = "", alias = ` alias="bu"`) =>
	`<link-entity name="businessunit" from="businessunitid" to="businessunitid"${alias}>${inner}</link-entity>`;

const warningsOf = (xml: string): ParseWarning[] => {
	const r = parseFetchXml(xml);
	expect(r.success, r.errors.map((e) => e.message).join("; ")).toBe(true);
	return r.warnings;
};
const errorsOf = (xml: string) => getExecutionErrors({ warnings: warningsOf(xml) });

describe("accepted by Dataverse → no errors", () => {
	it.each([
		["aggregate: group by + count + order by alias", SU(`${GRP}${CNT}<order alias="n" descending="true" />`, ` aggregate="true"`)],
		["aggregate: root order with entityname uses the link's alias", SU(`${CNT}${BU(`<attribute name="name" alias="g" groupby="true" />`)}<order entityname="bu" alias="g" />`, ` aggregate="true"`)],
		["aggregate: order inside the link uses the link's alias", SU(`${CNT}${BU(`<attribute name="name" alias="g" groupby="true" /><order alias="g" />`)}`, ` aggregate="true"`)],
		["aggregate: date grouping + UTC", SU(`<attribute name="createdon" alias="m" groupby="true" dategrouping="day" usertimezone="false" />${CNT}`, ` aggregate="true"`)],
		["aggregate: aggregatelimit without an order", SU(`${GRP}${CNT}`, ` aggregate="true" aggregatelimit="1000"`)],
		["aggregate: all-attributes is accepted", SU(`<all-attributes />${CNT}`, ` aggregate="true"`)],
		["non-aggregate: aggregatelimit + order (ignored)", SU(`<attribute name="fullname" /><order attribute="fullname" />`, ` top="3" aggregatelimit="100"`)],
		["non-aggregate: order entityname = link alias", SU(`<attribute name="fullname" />${BU(`<attribute name="name" />`)}<order entityname="bu" attribute="name" />`)],
		["non-aggregate: order entityname = link name", SU(`<attribute name="fullname" />${BU(`<attribute name="name" />`, "")}<order entityname="businessunit" attribute="name" />`)],
		["non-aggregate: order entityname = root entity name", SU(`<attribute name="fullname" /><order entityname="systemuser" attribute="fullname" />`)],
		["row aggregate on the root primary key with an alias", `<fetch><entity name="account"><attribute name="name" /><attribute name="accountid" rowaggregate="CountChildren" alias="kids" /></entity></fetch>`],
		["row aggregate next to the same attribute without alias (Learn)", `<fetch><entity name="account"><attribute name="accountid" /><attribute name="accountid" rowaggregate="CountChildren" alias="AccountChildren" /></entity></fetch>`],
		["union hint on an or-filter", SU(`<attribute name="fullname" /><filter type="or" hint="union"><condition attribute="isdisabled" operator="eq" value="0" /><condition attribute="isdisabled" operator="eq" value="1" /></filter>`)],
		["quick find with one override", SU(`<attribute name="fullname" /><filter type="or" isquickfindfields="1" overridequickfindrecordlimitdisabled="1"><condition attribute="fullname" operator="like" value="a%" /></filter>`)],
		["count + page + returntotalrecordcount (paging without top)", SU(`<attribute name="fullname" /><order attribute="fullname" />`, ` count="3" page="2" returntotalrecordcount="true"`)],
		["top alone", SU(`<attribute name="fullname" />`, ` top="5"`)],
		["alias equal to a link-entity alias", SU(`<attribute name="fullname" alias="bu" />${BU(`<attribute name="name" />`)}`)],
	])("%s", (_label, xml) => {
		expect(errorsOf(xml)).toEqual([]);
	});
});

describe("rejected by Dataverse → error", () => {
	it.each([
		["aggregate: attribute with no role (0x8004111e)", SU(`<attribute name="fullname" alias="f" />${CNT}`, ` aggregate="true"`), "0x8004111e"],
		["aggregate: link attribute with no role (0x8004111e)", SU(`${CNT}${BU(`<attribute name="name" alias="g" />`)}`, ` aggregate="true"`), "0x8004111e"],
		["aggregate: aggregate attribute without alias (0x8004110b)", SU(`<attribute name="systemuserid" aggregate="count" />`, ` aggregate="true"`), "0x8004110b"],
		["aggregate: link group by without alias (0x8004110b)", SU(`${CNT}${BU(`<attribute name="name" groupby="true" />`)}`, ` aggregate="true"`), "0x8004110b"],
		["aggregate: group by and aggregate on one attribute (0x80041137)", SU(`<attribute name="isdisabled" alias="d" groupby="true" aggregate="count" />`, ` aggregate="true"`), "0x80041137"],
		["non-aggregate: aggregate attribute (0x8004111a)", SU(`<attribute name="systemuserid" alias="n" aggregate="count" />`), "0x8004111a"],
		["non-aggregate: group by attribute (0x8004112e)", SU(`<attribute name="isdisabled" alias="d" groupby="true" />`), "0x8004112e"],
		["aggregate: root order alias defined in a link (0x8004110a)", SU(`${CNT}${BU(`<attribute name="name" alias="g" groupby="true" />`)}<order alias="g" />`, ` aggregate="true"`), "0x8004110a"],
		["aggregate: link order alias defined on the root (0x8004110a)", SU(`${GRP}${CNT}${BU(`<attribute name="name" alias="g" groupby="true" /><order alias="d" />`)}`, ` aggregate="true"`), "0x8004110a"],
		["aggregate: order alias that doesn't exist (0x8004110a)", SU(`${GRP}${CNT}<order alias="zzz" />`, ` aggregate="true"`), "0x8004110a"],
		["aggregate: order by attribute (0x80041131)", SU(`${GRP}${CNT}<order attribute="isdisabled" />`, ` aggregate="true"`), "0x80041131"],
		["non-aggregate: order by alias (0x80041132)", SU(`<attribute name="fullname" /><order attribute="fullname" alias="x" />`), "0x80041132"],
		["order entityname matching nothing (0x80041108)", SU(`<attribute name="fullname" /><order entityname="nope" attribute="fullname" />`), "0x80041108"],
		["entityname on an order inside a link (0x80041108)", SU(`<attribute name="fullname" />${BU(`<attribute name="name" /><order entityname="bu" attribute="name" />`)}`), "0x80041108"],
		["aggregate: aggregatelimit + order (SQL error)", SU(`${GRP}${CNT}<order alias="d" />`, ` aggregate="true" aggregatelimit="1000"`), "SQL error"],
		["union hint on an and-filter (0x80041122)", SU(`<attribute name="fullname" /><filter type="and" hint="union"><condition attribute="isdisabled" operator="eq" value="0" /></filter>`), "0x80041122"],
		["two union hints (0x80041122)", SU(`<attribute name="fullname" /><filter><filter type="or" hint="union"><condition attribute="isdisabled" operator="eq" value="0" /></filter><filter type="or" hint="union"><condition attribute="isdisabled" operator="eq" value="1" /></filter></filter>`), "0x80041122"],
		["hint other than 'union' (0x80041122)", SU(`<attribute name="fullname" /><filter type="or" hint="Union"><condition attribute="isdisabled" operator="eq" value="0" /></filter>`), "0x80041122"],
		["row aggregate without alias (SQL error)", `<fetch><entity name="account"><attribute name="accountid" rowaggregate="CountChildren" /></entity></fetch>`, "alias"],
		["row aggregate inside a link-entity", `<fetch><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="accountid" to="parentcustomerid" alias="a"><attribute name="accountid" rowaggregate="CountChildren" alias="kids" /></link-entity></entity></fetch>`, "root entity"],
		["row aggregate in an aggregate query (SQL error)", SU(`<attribute name="systemuserid" alias="k" groupby="true" rowaggregate="CountChildren" />${CNT.replace('alias="n"', 'alias="n2"')}`, ` aggregate="true"`), "aggregate query"],
		["duplicate alias (0x80041130)", SU(`<attribute name="fullname" alias="x" /><attribute name="domainname" alias="x" />`), "0x80041130"],
		["alias differing only in case (0x80041130)", SU(`<attribute name="fullname" alias="x" /><attribute name="domainname" alias="X" />`), "0x80041130"],
		["same alias on root and link attributes (0x80041130)", SU(`<attribute name="fullname" alias="x" />${BU(`<attribute name="name" alias="x" />`)}`), "0x80041130"],
		["alias equal to another root attribute's logical name (0x80041130)", SU(`<attribute name="fullname" alias="domainname" /><attribute name="domainname" />`), "0x80041130"],
		["duplicate link-entity alias (0x80041130)", SU(`<attribute name="fullname" />${BU(`<attribute name="name" />`)}${BU(`<attribute name="name" />`)}`), "0x80041130"],
		["top with count (0x80040203)", SU(`<attribute name="fullname" />`, ` top="5" count="3"`), "0x80040203"],
		["top with page (0x80040203)", SU(`<attribute name="fullname" />`, ` top="5" page="2"`), "0x80040203"],
		["top with returntotalrecordcount (0x80040203)", SU(`<attribute name="fullname" />`, ` top="5" returntotalrecordcount="true"`), "0x80040203"],
		["page 0 (0x80040203)", SU(`<attribute name="fullname" />`, ` count="3" page="0"`), "Page number"],
		["descending='True' (0x80041119)", SU(`<attribute name="fullname" /><order attribute="fullname" descending="True" />`), "0x80041119"],
	])("%s", (_label, xml, needle) => {
		const errors = errorsOf(xml);
		expect(errors.length, "expected an error").toBeGreaterThan(0);
		expect(errors.map((e) => e.message).join(" | ")).toContain(needle);
	});

	it("query-level findings point at the node they're about", () => {
		const r = parseFetchXml(SU(`<attribute name="fullname" alias="f" />${CNT}`, ` aggregate="true"`));
		const finding = getExecutionErrors(r).find((w) => w.message.includes("0x8004111e"))!;
		const node = r.fetchNode!.entity.attributes.find((a) => a.name === "fullname")!;
		expect(finding.nodeId).toBe(node.id);
	});
});

describe("accepted but ignored → info only", () => {
	it.each([
		["aggregatelimit on a non-aggregate query", SU(`<attribute name="fullname" />`, ` aggregatelimit="5"`), "aggregatelimit"],
		["date grouping outside an aggregate query", SU(`<attribute name="createdon" dategrouping="year" alias="y" />`), "dategrouping"],
		["date grouping on an aggregate attribute", SU(`<attribute name="createdon" aggregate="max" dategrouping="year" alias="y" />`, ` aggregate="true"`), "dategrouping"],
		["distinct with an aggregate other than countcolumn", SU(`<attribute name="systemuserid" alias="c" aggregate="count" distinct="true" />`, ` aggregate="true"`), "distinct"],
		["both quick find overrides", SU(`<attribute name="fullname" /><filter type="or" isquickfindfields="1" overridequickfindrecordlimitenabled="1" overridequickfindrecordlimitdisabled="1"><condition attribute="fullname" operator="like" value="%" /></filter>`), "overridequickfindrecordlimitenabled"],
		["aggregate='1' isn't an aggregate query", SU(`<attribute name="fullname" />`, ` aggregate="1"`), "aggregate"],
	])("%s", (_label, xml, attribute) => {
		const warnings = warningsOf(xml);
		expect(getExecutionErrors({ warnings })).toEqual([]);
		expect(warnings.some((w) => w.severity === "info" && w.attribute === attribute)).toBe(true);
	});
});

describe("condition aggregate (verified ignored: no HAVING in FetchXML)", () => {
	it("is kept, reported as ignored on the condition node, and never blocks Execute", () => {
		const xml = SU(`<attribute name="businessunitid" alias="b" groupby="true" />${CNT}<filter><condition attribute="systemuserid" aggregate="count" operator="ne" value="11111111-1111-1111-1111-111111111111" /></filter>`, ` aggregate="true"`);
		const parsed = parseFetchXml(xml);
		const condition = parsed.fetchNode!.entity.filters[0].conditions[0];
		expect(condition.aggregate).toBe("count");
		expect(getExecutionErrors(parsed)).toEqual([]);
		const finding = validateQueryRules(parsed.fetchNode!).find((f) => f.element === "condition");
		expect(finding).toMatchObject({ attribute: "aggregate", severity: "info", kind: "ignored", nodeId: condition.id });
	});
});

describe("valueof (filter-rows limitations, verified live 2026-10-14)", () => {
	const CT = (filter: string, links = "") =>
		`<fetch top="1"><entity name="contact"><attribute name="fullname" /><filter>${filter}</filter>${links}</entity></fetch>`;
	const ACCT = (inner = "", alias = ` alias="acct"`) =>
		`<link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer"${alias}>${inner}</link-entity>`;
	const valueofFindings = (xml: string) =>
		validateQueryRules(parseFetchXml(xml).fetchNode!).filter((f) => f.attribute === "valueof");

	it.each([
		["eq in the same row", CT(`<condition attribute="firstname" operator="eq" valueof="lastname" />`)],
		["ge on dates", CT(`<condition attribute="modifiedon" operator="ge" valueof="createdon" />`)],
		["a joined link-entity alias", CT(`<condition attribute="fullname" operator="eq" valueof="acct.name" />`, ACCT())],
		["a nested link-entity alias", CT(`<condition attribute="fullname" operator="eq" valueof="pc.fullname" />`, ACCT(`<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="pc" />`))],
		["the root entity's name", CT(`<condition attribute="firstname" operator="eq" valueof="contact.lastname" />`)],
	])("accepted: %s", (_label, xml) => {
		expect(valueofFindings(xml)).toEqual([]);
	});

	it.each([
		["like (0x80048415)", CT(`<condition attribute="firstname" operator="like" valueof="lastname" />`)],
		["begins-with (0x80048415)", CT(`<condition attribute="firstname" operator="begins-with" valueof="lastname" />`)],
		["on-or-after (0x80048415)", CT(`<condition attribute="modifiedon" operator="on-or-after" valueof="createdon" />`)],
		["a link-entity name without alias (0x80040216)", CT(`<condition attribute="fullname" operator="eq" valueof="account.name" />`, ACCT("", ""))],
		["an unknown qualifier (0x80040216)", CT(`<condition attribute="firstname" operator="eq" valueof="zz.lastname" />`)],
		["an alias in a different case (0x80040216)", CT(`<condition attribute="fullname" operator="eq" valueof="ACCT.name" />`, ACCT())],
		["the alias of a link-entity inside a filter (0x80040216)", CT(`<condition attribute="fullname" operator="eq" valueof="x.name" /><link-entity name="account" from="primarycontactid" to="contactid" link-type="any" alias="x" />`)],
	])("rejected: %s", (_label, xml) => {
		const findings = valueofFindings(xml);
		expect(findings.some((f) => f.severity === "error")).toBe(true);
	});

	describe("reach by where the condition is (verified live 2026-10-15)", () => {
		const c = (attr: string, valueof: string) => `<filter><condition attribute="${attr}" operator="eq" valueof="${valueof}" /></filter>`;
		const PC = (f = "") => `<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="pc">${f}</link-entity>`;
		const ACCT = (f = "", child = "") => `<link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer" alias="acct">${f}${child}</link-entity>`;
		const OWN = (f = "") => `<link-entity name="systemuser" from="systemuserid" to="owninguser" link-type="outer" alias="own">${f}</link-entity>`;
		const Q = (links: string, rootFilter = "") => `<fetch top="3"><entity name="contact"><attribute name="fullname" />${rootFilter}${links}</entity></fetch>`;
		const anyLink = (type: string, inner: string) => `<filter><link-entity name="account" from="primarycontactid" to="contactid" link-type="${type}" alias="x">${inner}</link-entity></filter>`;

		it.each([
			["root filter → nested alias", Q(ACCT("", PC()) + OWN(), c("fullname", "pc.fullname"))],
			["root filter, entityname acct → root name", Q(ACCT("", PC()), `<filter><condition entityname="acct" attribute="name" operator="eq" valueof="contact.fullname" /></filter>`)],
			["root filter, entityname acct → other alias", Q(ACCT() + OWN(), `<filter><condition entityname="acct" attribute="name" operator="eq" valueof="own.fullname" /></filter>`)],
			["link filter → its own alias", Q(ACCT(c("name", "acct.accountnumber")))],
			["link filter → root name", Q(ACCT(c("name", "contact.fullname")))],
			["nested link filter → parent alias", Q(ACCT("", PC(c("fullname", "acct.name"))))],
			["link filter → earlier sibling", Q(ACCT() + OWN(c("fullname", "acct.name")))],
			["link filter → descendant of an earlier sibling", Q(ACCT("", PC()) + OWN(c("fullname", "pc.fullname")))],
			["any filter in the root → any joined alias", Q(ACCT() + OWN(), anyLink("any", c("name", "own.fullname")))],
			["all filter in the root → root name", Q(ACCT(), anyLink("all", c("name", "contact.fullname")))],
			["any filter inside acct's filter → acct", Q(ACCT(`<filter><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="any">${c("fullname", "acct.name")}</link-entity></filter>`) + OWN())],
		])("accepted: %s", (_label, xml) => expect(valueofFindings(xml)).toEqual([]));

		it.each([
			["link filter → its child", Q(ACCT(c("name", "pc.fullname"), PC()))],
			["link filter → later sibling", Q(ACCT(c("name", "own.fullname")) + OWN())],
			["nested link filter → later uncle", Q(ACCT("", PC(c("fullname", "own.fullname"))) + OWN())],
			["any filter inside acct's filter → later alias", Q(ACCT(`<filter><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="any">${c("fullname", "own.fullname")}</link-entity></filter>`) + OWN())],
			["filter link's own alias", Q("", anyLink("any", c("name", "x.accountnumber")))],
		])("rejected (0x80040216): %s", (_label, xml) => {
			expect(valueofFindings(xml).some((f) => f.severity === "error")).toBe(true);
		});
	});

	it("is ignored by operators that take no value (null + valueof ran fine)", () => {
		const findings = valueofFindings(CT(`<condition attribute="firstname" operator="null" valueof="zz.lastname" />`));
		expect(findings).toHaveLength(1);
		expect(findings[0]).toMatchObject({ severity: "info", kind: "ignored" });
	});
});

describe("link-type by placement (verified live 2026-10-14)", () => {
	const F = `<filter><condition attribute="name" operator="not-null" /></filter>`;
	const JOINED = (lt: string, inner = F) =>
		`<fetch top="3"><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="primarycontactid" to="contactid" link-type="${lt}" alias="a">${inner}</link-entity></entity></fetch>`;
	const IN_FILTER = (lt: string, inner = F) =>
		`<fetch top="3"><entity name="contact"><attribute name="fullname" /><filter><link-entity name="account" from="primarycontactid" to="contactid" link-type="${lt}" alias="a">${inner}</link-entity></filter></entity></fetch>`;
	const linkErrors = (xml: string) =>
		validateQueryRules(parseFetchXml(xml).fetchNode!).filter((f) => f.element === "link-entity" && f.severity === "error");

	it.each(["inner", "outer", "exists", "in", "matchfirstrowusingcrossapply", "any", "not any", "all", "not all"])(
		"joined %s is accepted",
		(lt) => expect(linkErrors(JOINED(lt))).toEqual([]),
	);
	it.each(["any", "not any", "all", "not all"])("%s inside a filter is accepted", (lt) =>
		expect(linkErrors(IN_FILTER(lt))).toEqual([]),
	);
	it.each(["inner", "outer", "exists", "in", "matchfirstrowusingcrossapply"])(
		"%s inside a filter is rejected (0x8004111c)",
		(lt) => expect(linkErrors(IN_FILTER(lt))).toHaveLength(1),
	);
	it.each([
		["exists + attribute", JOINED("exists", `<attribute name="name" />`)],
		["in + attribute", JOINED("in", `<attribute name="name" />`)],
		["any + all-attributes", JOINED("any", `<all-attributes />`)],
		["not any + order", IN_FILTER("not any", `<order attribute="name" />${F}`)],
		["all + attribute", IN_FILTER("all", `<attribute name="name" />${F}`)],
	])("%s is rejected (0x80041121)", (_label, xml) => expect(linkErrors(xml)).toHaveLength(1));
	it.each([
		["inner + attribute", JOINED("inner", `<attribute name="name" />`)],
		["outer + attribute", JOINED("outer", `<attribute name="name" />`)],
		["matchfirstrowusingcrossapply + attribute", JOINED("matchfirstrowusingcrossapply", `<attribute name="name" />`)],
		["exists + nested link-entity", JOINED("exists", `<link-entity name="account" from="accountid" to="parentcustomerid" alias="b" />`)],
	])("%s is accepted", (_label, xml) => expect(linkErrors(xml)).toEqual([]));
	it("visible is reported as ignored, never an error", () => {
		const findings = validateQueryRules(parseFetchXml(JOINED("inner", `<attribute name="name" />`).replace(`alias="a"`, `alias="a" visible="false"`)).fetchNode!);
		expect(findings.find((f) => f.attribute === "visible")).toMatchObject({ severity: "info", kind: "ignored" });
	});
});

describe("more than one filter in an entity (#94; verified: combined with and)", () => {
	it("is advice on the extra filter, never an error", () => {
		const parsed = parseFetchXml(SU(`<attribute name="fullname" /><filter><condition attribute="fullname" operator="not-null" /></filter><filter><condition attribute="isdisabled" operator="eq" value="0" /></filter>`));
		const fetch = parsed.fetchNode!;
		expect(fetch.entity.filters).toHaveLength(2);
		expect(getExecutionErrors(parsed)).toEqual([]);
		const findings = validateQueryRules(fetch).filter((f) => f.element === "filter");
		expect(findings).toEqual([expect.objectContaining({ severity: "info", kind: "advice", nodeId: fetch.entity.filters[1].id })]);
	});
});

describe("boolean spellings Dataverse accepts are read correctly", () => {
	const attrOf = (xml: string) => parseFetchXml(xml).fetchNode!.entity.attributes[0];
	it("usertimezone '1' / '0' keep their meaning (was inverted)", () => {
		expect(attrOf(SU(`<attribute name="createdon" alias="d" groupby="true" dategrouping="day" usertimezone="1" />${CNT}`, ` aggregate="true"`)).usertimezone).toBe(true);
		expect(attrOf(SU(`<attribute name="createdon" alias="d" groupby="true" dategrouping="day" usertimezone="0" />${CNT}`, ` aggregate="true"`)).usertimezone).toBe(false);
	});
	it("groupby / distinct / returntotalrecordcount accept '1' and any case", () => {
		expect(attrOf(SU(`<attribute name="isdisabled" alias="d" groupby="1" />${CNT}`, ` aggregate="true"`)).groupby).toBe(true);
		expect(attrOf(SU(`<attribute name="isdisabled" alias="d" groupby="True" />${CNT}`, ` aggregate="true"`)).groupby).toBe(true);
		const f = parseFetchXml(SU(`<attribute name="fullname" />`, ` distinct="1" returntotalrecordcount="TRUE"`)).fetchNode!;
		expect(f.options.distinct).toBe(true);
		expect(f.options.returnTotalRecordCount).toBe(true);
	});
	it("descending '1' sorts descending", () => {
		const o = parseFetchXml(SU(`<attribute name="fullname" /><order attribute="fullname" descending="1" />`)).fetchNode!.entity.orders[0];
		expect(o.descending).toBe(true);
	});
	it("enableprefiltering='1' (the documented form) is kept", () => {
		const e = parseFetchXml(`<fetch><entity name="account" enableprefiltering="1" prefilterparametername="CRM_FilteredAccount"><attribute name="name" /></entity></fetch>`).fetchNode!.entity;
		expect(e.enablePrefiltering).toBe(true);
	});
});

describe("condition value shapes (#87, verified live 2026-10-15)", () => {
	const C = (cond: string) => `<fetch top="1"><entity name="account"><attribute name="name" /><filter>${cond}</filter></entity></fetch>`;
	const valueFindings = (xml: string) =>
		validateQueryRules(parseFetchXml(xml).fetchNode!).filter((f) => f.element === "condition" && f.attribute === "value");
	it.each([
		["last-x-days 7", C(`<condition attribute="createdon" operator="last-x-days" value="7" />`)],
		["in-fiscal-year 2023", C(`<condition attribute="createdon" operator="in-fiscal-year" value="2023" />`)],
		["in-fiscal-period-and-year 3, 2023", C(`<condition attribute="createdon" operator="in-fiscal-period-and-year"><value>3</value><value>2023</value></condition>`)],
		["between two values", C(`<condition attribute="numberofemployees" operator="between"><value>1</value><value>9</value></condition>`)],
		["in one value", C(`<condition attribute="name" operator="in"><value>a</value></condition>`)],
		["eq value", C(`<condition attribute="name" operator="eq" value="a" />`)],
		["eq valueof", C(`<condition attribute="name" operator="eq" valueof="accountnumber" />`)],
	])("accepted: %s", (_l, xml) => expect(valueFindings(xml)).toEqual([]));
	it.each([
		["last-x-days ISO date (0x80041139)", C(`<condition attribute="createdon" operator="last-x-days" value="2026-10-12T00:00:00.000Z" />`)],
		["last-x-days 0 (0x80040203)", C(`<condition attribute="createdon" operator="last-x-days" value="0" />`)],
		["last-x-days 1.5 (0x80041139)", C(`<condition attribute="createdon" operator="last-x-days" value="1.5" />`)],
		["last-x-days as <value> (0x8004111e)", C(`<condition attribute="createdon" operator="last-x-days"><value>7</value></condition>`)],
		["in-fiscal-year 23 (0x80040203)", C(`<condition attribute="createdon" operator="in-fiscal-year" value="23" />`)],
		["in-fiscal-period-and-year one value (0x80041106)", C(`<condition attribute="createdon" operator="in-fiscal-period-and-year"><value>3</value></condition>`)],
		["in-fiscal-period-and-year dates", C(`<condition attribute="createdon" operator="in-fiscal-period-and-year"><value>2024-01-01</value><value>2024-12-31</value></condition>`)],
		["between one value (0x80041106)", C(`<condition attribute="numberofemployees" operator="between"><value>1</value></condition>`)],
		["in value attribute (0x80040203)", C(`<condition attribute="name" operator="in" value="a" />`)],
		["eq with <value> elements (0x8004111e)", C(`<condition attribute="name" operator="eq"><value>a</value><value>b</value></condition>`)],
		["eq without value (0x8004111e)", C(`<condition attribute="name" operator="eq" />`)],
	])("rejected: %s", (_l, xml) => expect(valueFindings(xml).some((f) => f.severity === "error")).toBe(true));
	it("a value on an operator that takes none is a note (accepted, ignored)", () => {
		expect(valueFindings(C(`<condition attribute="createdon" operator="today" value="1" />`))).toEqual([
			expect.objectContaining({ severity: "info" }),
		]);
	});
});
