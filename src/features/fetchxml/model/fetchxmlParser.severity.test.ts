/**
 * Parser severity classification (#89, #48).
 *
 * Principle: what the FetchXML reference documents is supported; undocumented behaviour verified
 * to work is supported too; only what is known not to work is blocked (error); what can't be
 * determined client-side is a warning with "Execute anyway".
 *
 * Each case mirrors a probe run against a live Dataverse org through the PPTB host request path
 * (2026-10-13). The comment on each case records what Dataverse did with the raw FetchXML.
 */
import { describe, expect, it } from "vitest";
import {
	getBlockingWarnings,
	getExecutionErrors,
	getParseToTreeChanges,
	parseFetchXml,
	type ParseWarning,
} from "./fetchxmlParser";

const wrap = (inner: string, fetchAttrs = "") =>
	`<fetch top="1"${fetchAttrs}><entity name="account"><attribute name="name" />${inner}</entity></fetch>`;
const withAttribute = (attributeXml: string) =>
	wrap("").replace('<attribute name="name" />', attributeXml);

function findings(xml: string) {
	const result = parseFetchXml(xml);
	return {
		result,
		errors: getExecutionErrors(result),
		warnings: getBlockingWarnings(result),
		changes: getParseToTreeChanges(result),
		only: (pred: (w: ParseWarning) => boolean) => result.warnings.filter(pred),
	};
}

describe("supported: documented, or undocumented but verified to work", () => {
	it("#48: xmlns:generator is ignored, never blocks, and changes nothing on Parse to Tree", () => {
		// Dataverse: ✅ runs
		const f = findings(`<fetch xmlns:generator="MarkMpn.SQL4CDS"><entity name="account"><attribute name="name" /></entity></fetch>`);
		expect(f.result.success).toBe(true);
		expect(f.result.warnings).toEqual([
			expect.objectContaining({ attribute: "xmlns:generator", severity: "info", kind: "ignored" }),
		]);
		expect(f.errors).toEqual([]);
		expect(f.warnings).toEqual([]);
		expect(f.changes).toEqual([]);
	});

	it.each([
		["fetch", `<fetch top="1" foo="1"><entity name="account"><attribute name="name" /></entity></fetch>`],
		["entity", `<fetch top="1"><entity name="account" bogus="x"><attribute name="name" /></entity></fetch>`],
		["attribute", withAttribute('<attribute name="name" bogus="x" />')],
		["filter", wrap(`<filter type="and" bogus="x"><condition attribute="statecode" operator="eq" value="0" /></filter>`)],
		["link-entity", wrap(`<link-entity name="contact" from="parentcustomerid" to="accountid" alias="c" bogus="x" />`)],
	])("unknown attribute on <%s> is a note (Dataverse ignores it)", (element, xml) => {
		const f = findings(xml);
		expect(f.result.success).toBe(true);
		expect(f.only((w) => w.element === element)).toEqual([
			expect.objectContaining({ attribute: expect.stringMatching(/^(foo|bogus)$/), severity: "info", kind: "ignored" }),
		]);
		expect(f.errors).toEqual([]);
		expect(f.warnings).toEqual([]);
		expect(f.changes).toEqual([]);
	});

	it("namespaced attribute is a note (Dataverse ignores it)", () => {
		const f = findings(`<fetch top="1" xmlns:x="urn:x"><entity name="account" x:note="hi"><attribute name="name" /></entity></fetch>`);
		expect(f.only((w) => w.attribute === "x:note")).toEqual([
			expect.objectContaining({ severity: "info", kind: "ignored" }),
		]);
		expect(f.errors).toEqual([]);
	});

	it("legacy Advanced Find / saved-view attributes are accepted silently", () => {
		// Dataverse: ✅ ignores version/output-format/mapping and condition uiname/uitype
		const f = findings(
			`<fetch version="1.0" output-format="xml-platform" mapping="logical" top="1"><entity name="account"><attribute name="name" />` +
				`<filter type="and"><condition attribute="ownerid" operator="ne" value="{00000000-0000-0000-0000-000000000001}" uiname="x" uitype="systemuser" /></filter>` +
				`</entity></fetch>`,
		);
		expect(f.result.warnings).toEqual([]);
	});

	it.each([
		["useraworderby", ` useraworderby="1"`],
		["latematerialize", ` latematerialize="true"`],
		["options", ` options="HashJoin"`],
		["datasource (retained)", ` datasource="retained"`],
		["datasource (bin, accepted by Dataverse)", ` datasource="bin"`],
	])("documented '%s' is supported with no findings (#90)", (_label, fetchAttrs) => {
		// Dataverse: ✅ accepts (datasource errors only where long-term retention isn't provisioned)
		const f = findings(wrap("", fetchAttrs));
		expect(f.result.warnings).toEqual([]);
		expect(f.changes).toEqual([]);
	});

	it("documented 'aggregatelimit' on an aggregate query is supported with no findings (#90)", () => {
		// Dataverse: ✅ accepts; on a non-aggregate query it is ignored (reported as info)
		const f = findings(
			`<fetch aggregate="true" aggregatelimit="1000"><entity name="task"><attribute name="prioritycode" alias="p" groupby="true" /><attribute name="activityid" alias="n" aggregate="count" /></entity></fetch>`,
		);
		expect(f.result.warnings).toEqual([]);
		expect(f.changes).toEqual([]);
	});

	it("order by alias in an aggregate query is supported and kept (#90)", () => {
		// Dataverse: ✅ accepts <order alias> in aggregate queries
		const f = findings(
			`<fetch aggregate="true"><entity name="task"><attribute name="prioritycode" alias="p" groupby="true" /><attribute name="activityid" alias="n" aggregate="count" /><order alias="n" descending="true" /></entity></fetch>`,
		);
		expect(f.result.warnings).toEqual([]);
		expect(f.result.fetchNode!.entity.orders).toEqual([
			expect.objectContaining({ alias: "n", descending: true }),
		]);
	});

	it("undocumented query hint names get advice (Dataverse silently ignores them)", () => {
		const f = findings(wrap("", ` options="HashJoin,HashJion"`));
		expect(f.result.warnings).toEqual([
			expect.objectContaining({ attribute: "options", severity: "info", kind: "advice", message: expect.stringContaining("HashJion") }),
		]);
		expect(f.result.fetchNode!.options.queryHints).toEqual(["HashJoin", "HashJion"]);
	});

	it("link-type 'natural' (undocumented, verified to work) is supported and preserved", () => {
		const f = findings(wrap(`<link-entity name="contact" from="parentcustomerid" to="accountid" link-type="natural" alias="c" />`));
		expect(f.result.warnings).toEqual([]);
		expect(f.result.fetchNode!.entity.links[0].linkType).toBe("natural");
	});

	it("deprecated operator 'neq' is supported with advice", () => {
		// Dataverse: ✅ runs
		const f = findings(wrap(`<filter><condition attribute="statecode" operator="neq" value="1" /></filter>`));
		expect(f.result.warnings).toEqual([expect.objectContaining({ severity: "info", kind: "advice" })]);
	});

	it("documented fiscal 'x' operators are supported", () => {
		const f = findings(wrap(`<filter><condition attribute="createdon" operator="last-x-fiscal-years" value="2" /></filter>`));
		expect(f.result.warnings).toEqual([]);
	});

	it("top above 5,000 runs, with advice that Dataverse caps it", () => {
		// Dataverse: ✅ top="6000" returns 5,000 rows
		const f = findings(`<fetch top="6000"><entity name="account"><attribute name="name" /></entity></fetch>`);
		expect(f.result.warnings).toEqual([expect.objectContaining({ attribute: "top", severity: "info", kind: "advice" })]);
		expect(f.changes).toEqual([]);
	});

	it("non-integer utc-offset is a note (Dataverse ignores it)", () => {
		const f = findings(wrap("", ` utc-offset="abc"`));
		expect(f.result.warnings).toEqual([expect.objectContaining({ severity: "info", kind: "ignored" })]);
	});

	it("extra <entity> elements: Dataverse uses the first; Parse to Tree asks before dropping", () => {
		// Dataverse: ✅ runs
		const f = findings(`<fetch top="1"><entity name="account"><attribute name="name" /></entity><entity name="contact" /></fetch>`);
		expect(f.errors).toEqual([]);
		expect(f.changes).toEqual([expect.objectContaining({ element: "fetch", kind: "dropped" })]);
	});

	it("link-entity alias with a dot is a note (Dataverse accepts it; Parse to Tree tidies it)", () => {
		const f = findings(wrap(`<link-entity name="contact" from="parentcustomerid" to="accountid" alias="a.b" />`));
		expect(f.only((w) => w.element === "link-entity")).toEqual([
			expect.objectContaining({ attribute: "alias", severity: "info", kind: "corrected" }),
		]);
		expect(f.errors).toEqual([]);
		expect(f.changes).toEqual([]);
		expect(f.result.fetchNode!.entity.links[0].alias).toBe("ab");
	});
});

describe("errors: known to be rejected by Dataverse (Execute blocked)", () => {
	it.each([
		// [label, xml, expected kind, Dataverse error]
		["unknown child element", wrap("<foo />"), "dropped"], // 0x8004111c
		["namespaced child element", `<fetch top="1" xmlns:x="urn:x"><entity name="account"><attribute name="name" /><x:meta /></entity></fetch>`, "dropped"], // 0x8004111c
		["attribute alias 'my name'", withAttribute('<attribute name="name" alias="my name" />'), "corrected"], // 0x80041109
		["attribute alias 'my-name'", withAttribute('<attribute name="name" alias="my-name" />'), "corrected"], // 0x80041109
		["attribute alias '1name'", withAttribute('<attribute name="name" alias="1name" />'), "corrected"], // 0x80041109
		["non-integer top", `<fetch top="abc"><entity name="account"><attribute name="name" /></entity></fetch>`, "dropped"], // Input string was not in a correct format
		["unknown link-type", wrap(`<link-entity name="contact" from="parentcustomerid" to="accountid" link-type="bogus" alias="c" />`), "corrected"], // 0x80041117
		["attribute without name", withAttribute("<attribute />"), "invalid"], // 0x8004111e
		["condition without attribute", wrap(`<filter><condition operator="eq" value="x" /></filter>`), "invalid"], // 0x8004111e
		["condition without operator", wrap(`<filter><condition attribute="name" value="x" /></filter>`), "corrected"], // 0x8004111e
		["order with neither attribute nor alias", wrap(`<order descending="true" />`), "dropped"], // 0x80041133
		["default xmlns", `<fetch xmlns="urn:test"><entity name="account"><attribute name="name" /></entity></fetch>`, "corrected"], // 0x80041102
		["invalid filter type", wrap(`<filter type="xor"><condition attribute="statecode" operator="eq" value="0" /></filter>`), "corrected"], // 0x80041122
		["link-entity without name", wrap(`<link-entity from="parentcustomerid" to="accountid" alias="c" />`), "invalid"], // 0x8004111e
		["invalid aggregate", `<fetch aggregate="true"><entity name="account"><attribute name="accountid" alias="n" aggregate="median" /></entity></fetch>`, "dropped"], // 0x8004111a
		["invalid dategrouping", `<fetch aggregate="true"><entity name="account"><attribute name="createdon" alias="d" groupby="true" dategrouping="decade" /></entity></fetch>`, "dropped"], // 0x80041135
	])("%s", (_label, xml, kind) => {
		const f = findings(xml);
		expect(f.result.success).toBe(true); // Parse to Tree can still build a tree
		expect(f.errors).toEqual(expect.arrayContaining([expect.objectContaining({ severity: "error", kind })]));
	});

	it.each([
		["count", `<fetch top="5" count="5">`],
		["page", `<fetch top="5" page="2">`],
		["returntotalrecordcount", `<fetch top="5" returntotalrecordcount="true">`],
	])("top combined with %s is an error (0x80040203)", (_label, fetchTag) => {
		const f = findings(`${fetchTag}<entity name="account"><attribute name="name" /></entity></fetch>`);
		expect(f.errors).toEqual([expect.objectContaining({ attribute: "top", kind: "invalid" })]);
	});

	it.each([
		["aggregatelimit above 50,000", `<fetch aggregate="true" aggregatelimit="60000"><entity name="task"><attribute name="activityid" alias="n" aggregate="count" /></entity></fetch>`, "aggregatelimit", "dropped"], // "Expected value between 1 and 50000 inclusive."
		["aggregatelimit not a whole number", `<fetch aggregate="true" aggregatelimit="abc"><entity name="task"><attribute name="activityid" alias="n" aggregate="count" /></entity></fetch>`, "aggregatelimit", "dropped"],
		["datasource other than retained/bin", wrap("", ` datasource="live"`), "datasource", "invalid"], // 0x80041108
		["aggregate='rowaggregate' (use rowaggregate='CountChildren')", `<fetch aggregate="true"><entity name="account"><attribute name="accountid" alias="kids" aggregate="rowaggregate" /></entity></fetch>`, "aggregate", "dropped"], // 0x8004111a
		["order by attribute in an aggregate query", `<fetch aggregate="true"><entity name="task"><attribute name="prioritycode" alias="p" groupby="true" /><attribute name="activityid" alias="n" aggregate="count" /><order attribute="prioritycode" /></entity></fetch>`, "attribute", "invalid"], // 0x80041131
		["order by alias in a non-aggregate query", withAttribute('<attribute name="name" alias="nm" />').replace("</entity>", '<order alias="nm" /></entity>'), "alias", "invalid"], // 0x80041132
	])("%s is an error", (_label, xml, attribute, kind) => {
		const f = findings(xml);
		expect(f.errors).toEqual(expect.arrayContaining([expect.objectContaining({ attribute, kind })]));
	});

	it("order rules apply inside link-entities too", () => {
		const f = findings(
			`<fetch aggregate="true"><entity name="contact"><attribute name="contactid" alias="n" aggregate="count" /><link-entity name="account" from="accountid" to="parentcustomerid" alias="a"><attribute name="name" alias="acct" groupby="true" /><order attribute="name" /></link-entity></entity></fetch>`,
		);
		expect(f.errors).toEqual([expect.objectContaining({ element: "order", attribute: "attribute" })]);
	});

	it("errors that Parse to Tree corrects or drops are listed as changes", () => {
		const f = findings(withAttribute('<attribute name="name" alias="my name" />'));
		expect(f.changes).toEqual([expect.objectContaining({ attribute: "alias", kind: "corrected" })]);
	});
});

describe("warnings: can't be determined client-side (Execute anyway)", () => {
	it("link-entity without from/to (valid only for system many-to-many relationships)", () => {
		// Dataverse: ❌ contact→account (N:1), ✅ systemuser→team (system N:N)
		const f = findings(`<fetch top="1"><entity name="contact"><attribute name="fullname" /><link-entity name="account" alias="a" /></entity></fetch>`);
		expect(f.errors).toEqual([]);
		expect(f.warnings).toEqual([expect.objectContaining({ element: "link-entity", kind: "invalid" })]);
	});

	it("rowaggregate value other than the documented 'CountChildren'", () => {
		const f = findings(withAttribute('<attribute name="accountid" alias="k" rowaggregate="CountParents" />'));
		expect(f.warnings).toEqual([expect.objectContaining({ attribute: "rowaggregate" })]);
	});

	it("operator not in the documented list (Microsoft may add operators)", () => {
		// Dataverse: ❌ 0x80041120 for 'bogus'
		const f = findings(wrap(`<filter><condition attribute="name" operator="bogus" value="x" /></filter>`));
		expect(f.errors).toEqual([]);
		expect(f.warnings).toEqual([expect.objectContaining({ attribute: "operator", kind: "invalid" })]);
	});
});

describe("fatal: no usable tree", () => {
	it("link-entity with empty from and to (Dataverse silently returns 0 rows)", () => {
		const r = parseFetchXml(
			`<fetch top="2"><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="" to="" alias="a" /></entity></fetch>`,
		);
		expect(r.success).toBe(false);
		expect(r.errors[0].message).toMatch(/empty 'from' and 'to'/);
	});
});
