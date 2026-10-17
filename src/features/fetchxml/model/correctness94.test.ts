/**
 * #94 / #63 correctness fixes: serializer escaping, references to corrected aliases, the editor's
 * intellisense catalogue, and the code generators.
 */
import { describe, expect, it } from "vitest";
import type * as Monaco from "monaco-editor";
import { parseFetchXml } from "./fetchxmlParser";
import { generateFetchXml } from "./fetchxml";
import { FETCHXML_SCHEMA, getElementContext } from "./fetchxmlIntellisense";
import { KNOWN_OPERATORS } from "./operators";
import { generateFxsSyncCode, generatePowerAutomateSpec } from "../engine/fetchxmlCodeGenerators";

const roundTrip = (xml: string) => {
	const first = parseFetchXml(xml);
	expect(first.success, first.errors.map((e) => e.message).join("; ")).toBe(true);
	return generateFetchXml(first.fetchNode!);
};

describe("serializer escaping (#63, #94)", () => {
	it("keeps apostrophes literal in attribute values (Power Automate expressions)", () => {
		const out = roundTrip(
			`<fetch><entity name="account"><attribute name="name" /><filter><condition attribute="name" operator="eq" value="@{body('Parse_payload')?['Student']?['RecordID']}" /></filter></entity></fetch>`,
		);
		expect(out).toContain(`value="@{body('Parse_payload')?['Student']?['RecordID']}"`);
		expect(out).not.toContain("&apos;");
	});

	it("escapes quotes, ampersands and angle brackets, and round-trips them unchanged", () => {
		const xml = `<fetch><entity name="account"><attribute name="name" /><filter><condition attribute="name" operator="eq" value="a &quot;b&quot; &amp; &lt;c&gt; 'd'" /><condition attribute="name" operator="in"><value>x &amp; 'y' &lt;z&gt;</value></condition></filter></entity></fetch>`;
		const out = roundTrip(xml);
		expect(out).toContain(`value="a &quot;b&quot; &amp; &lt;c&gt; 'd'"`);
		expect(out).toContain(`<value>x &amp; 'y' &lt;z&gt;</value>`);
		expect(roundTrip(out)).toBe(out);
		const condition = parseFetchXml(out).fetchNode!.entity.filters[0].conditions[0];
		expect(condition.value).toBe(`a "b" & <c> 'd'`);
	});

	it("escapes names, aliases, from / to and order references too", () => {
		const fetch = parseFetchXml(
			`<fetch><entity name="account"><attribute name="name" alias="n" /><order attribute="name" /><link-entity name="contact" from="contactid" to="primarycontactid" alias="pc" /></entity></fetch>`,
		).fetchNode!;
		fetch.entity.attributes[0].name = `na"me`;
		fetch.entity.orders[0].attribute = `na"me`;
		fetch.entity.links[0].from = `a<b`;
		fetch.entity.links[0].to = `c&d`;
		const out = generateFetchXml(fetch);
		expect(out).toContain(`name="na&quot;me"`);
		expect(out).toContain(`attribute="na&quot;me"`);
		expect(out).toContain(`from="a&lt;b"`);
		expect(out).toContain(`to="c&amp;d"`);
		expect(parseFetchXml(out).success).toBe(true);
	});
});

describe("references to corrected aliases follow the correction (#94)", () => {
	it("link-entity alias: condition entityname, valueof qualifier and order entityname", () => {
		const result = parseFetchXml(
			`<fetch><entity name="contact"><attribute name="fullname" /><order entityname="my.acct" attribute="name" /><filter><condition entityname="my.acct" attribute="name" operator="null" /><condition attribute="fullname" operator="eq" valueof="my.acct.name" /></filter><link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer" alias="my.acct"><attribute name="name" /></link-entity></entity></fetch>`,
		);
		const fetch = result.fetchNode!;
		const newAlias = fetch.entity.links[0].alias!;
		expect(newAlias).not.toBe("my.acct");
		expect(fetch.entity.orders[0].entityname).toBe(newAlias);
		expect(fetch.entity.filters[0].conditions[0].entityname).toBe(newAlias);
		expect(fetch.entity.filters[0].conditions[1].valueof).toBe(`${newAlias}.name`);
		expect(result.warnings.some((w) => w.kind === "corrected" && w.message.includes("references"))).toBe(true);
	});

	it("attribute alias: aggregate order alias", () => {
		const fetch = parseFetchXml(
			`<fetch aggregate="true"><entity name="account"><attribute name="accountid" alias="row count" aggregate="count" /><order alias="row count" /></entity></fetch>`,
		).fetchNode!;
		expect(fetch.entity.orders[0].alias).toBe(fetch.entity.attributes[0].alias);
	});

	it("leaves queries without corrections alone", () => {
		const result = parseFetchXml(
			`<fetch><entity name="contact"><attribute name="fullname" /><filter><condition entityname="acct" attribute="name" operator="null" /></filter><link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer" alias="acct" /></entity></fetch>`,
		);
		expect(result.fetchNode!.entity.filters[0].conditions[0].entityname).toBe("acct");
		expect(result.warnings.some((w) => w.message.includes("references"))).toBe(false);
	});
});

describe("intellisense catalogue matches the FetchXML reference (#94)", () => {
	const attrs = (element: string) => Object.keys(FETCHXML_SCHEMA.elements[element].attributes);
	it.each([
		// Every documented attribute (learn.microsoft.com/.../fetchxml/reference/*)
		["fetch", ["aggregate", "aggregatelimit", "count", "datasource", "distinct", "latematerialize", "no-lock", "options", "page", "paging-cookie", "returntotalrecordcount", "top", "useraworderby"]],
		["entity", ["name"]],
		["link-entity", ["name", "to", "from", "alias", "link-type", "intersect"]],
		["attribute", ["aggregate", "alias", "dategrouping", "distinct", "groupby", "name", "rowaggregate", "usertimezone"]],
		["order", ["attribute", "alias", "descending", "entityname"]],
		["filter", ["type", "hint", "isquickfindfields", "overridequickfindrecordlimitenabled", "overridequickfindrecordlimitdisabled"]],
		["condition", ["attribute", "entityname", "operator", "value", "valueof"]],
	])("%s lists every documented attribute", (element, documented) => {
		expect(attrs(element)).toEqual(expect.arrayContaining(documented));
	});

	it("doesn't offer condition aggregate (Dataverse ignores it) or order under fetch", () => {
		expect(attrs("condition")).not.toContain("aggregate");
		expect(FETCHXML_SCHEMA.elements.fetch.children).toEqual(["entity"]);
	});

	it("offers every operator the parser accepts", () => {
		expect(FETCHXML_SCHEMA.elements.condition.attributes.operator.values).toEqual([...KNOWN_OPERATORS]);
	});

	it("suggests values for hyphenated attributes (link-type, no-lock)", () => {
		const line = `<link-entity name="contact" link-type="`;
		const model = {
			getValueInRange: () => `<fetch><entity name="account">${line}`,
			getLineContent: () => line,
		} as unknown as Monaco.editor.ITextModel;
		const context = getElementContext(model, { lineNumber: 1, column: line.length + 1 } as Monaco.Position);
		expect(context).toMatchObject({ inValue: true, attributeName: "link-type" });
	});
});

describe("code generators (#94)", () => {
	const pa = (xml: string) => generatePowerAutomateSpec(xml, "accounts").fields.find((f) => f.label === "Top Count")?.value;
	it("Power Automate Top Count comes from top only, never from count (a page size)", () => {
		expect(pa(`<fetch top="25"><entity name="account" /></fetch>`)).toBe("25");
		expect(pa(`<fetch count="50" page="1"><entity name="account" /></fetch>`)).toBe("");
		expect(pa(`<fetch><entity name="account"><attribute name="accountid" alias="top" aggregate="count" /></entity></fetch>`)).toBe("");
	});

	it("Web API snippet asks for annotations and explains paging cookies", () => {
		const code = generateFxsSyncCode("webapi", `<fetch count="50" page="1"><entity name="account" /></fetch>`, "accounts");
		expect(code).toContain(`Prefer: odata.include-annotations="*"`);
		expect(code).toContain("@Microsoft.Dynamics.CRM.fetchxmlpagingcookie");
		expect(code).toContain("@Microsoft.Dynamics.CRM.morerecords");
	});
});
