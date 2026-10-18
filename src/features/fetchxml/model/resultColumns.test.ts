/**
 * Results grid columns from the query, not from row 1 (#91). Row shapes mirror the Web API: null values
 * are left out, root lookups are _name_value, cross-apply columns use schema names.
 */
import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { orderByLayout, resultColumns, rowKeys, type ResultAttributeLookup } from "./resultColumns";
import { planSortByColumn, resolveColumn } from "./queryColumns";
import type { LayoutXmlConfig } from "./layoutxml";

const parse = (xml: string) => parseFetchXml(xml).fetchNode!;
const keys = (...args: Parameters<typeof resultColumns>) => resultColumns(...args).map((c) => c.key);
const meta =
	(types: Record<string, { AttributeType?: string; SchemaName?: string }>): ResultAttributeLookup =>
	(entity, attribute) =>
		types[`${entity}.${attribute}`];

const ACCOUNT = `<fetch><entity name="account"><attribute name="name" /><attribute name="primarycontactid" /><attribute name="address1_city" /></entity></fetch>`;

describe("resultColumns: the query's attributes", () => {
	it("a lookup null in row 1 keeps its _name_value column (found in a later row)", () => {
		const rows = [{ accountid: "1", name: "A" }, { accountid: "2", name: "B", _primarycontactid_value: "c" }];
		expect(keys(parse(ACCOUNT), rows)).toEqual(["name", "_primarycontactid_value", "address1_city"]);
	});

	it("metadata decides the lookup key even when no loaded row has a value", () => {
		const types = meta({ "account.primarycontactid": { AttributeType: "Lookup" }, "account.name": { AttributeType: "String" } });
		expect(keys(parse(ACCOUNT), [{ name: "A" }], types)).toEqual(["name", "_primarycontactid_value", "address1_city"]);
	});

	it.each(["Customer", "Owner"])("%s attributes are lookups too", (type) => {
		const fetch = parse(`<fetch><entity name="contact"><attribute name="parentcustomerid" /></entity></fetch>`);
		expect(keys(fetch, [], meta({ "contact.parentcustomerid": { AttributeType: type } }))).toEqual(["_parentcustomerid_value"]);
	});

	it("aliased lookups and link-entity lookups keep their key", () => {
		const fetch = parse(`<fetch><entity name="account"><attribute name="primarycontactid" alias="pc" /><link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="c"><attribute name="parentcustomerid" /></link-entity></entity></fetch>`);
		const types = meta({ "account.primarycontactid": { AttributeType: "Lookup" }, "contact.parentcustomerid": { AttributeType: "Customer" } });
		expect(keys(fetch, [], types)).toEqual(["pc", "c.parentcustomerid"]);
	});

	it("unaliased link-entities use {name}{N}", () => {
		const fetch = parse(`<fetch><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer"><attribute name="name" /></link-entity></entity></fetch>`);
		expect(keys(fetch, [{ fullname: "F", "account1.name": "A" }])).toEqual(["fullname", "account1.name"]);
	});

	it("exists and in link-entities return no columns, nor do link-entities under them", () => {
		const fetch = parse(`<fetch><entity name="contact"><attribute name="fullname" /><link-entity name="account" from="primarycontactid" to="contactid" link-type="exists"><attribute name="name" /><link-entity name="systemuser" from="systemuserid" to="owninguser"><attribute name="fullname" /></link-entity></link-entity><link-entity name="account" from="primarycontactid" to="contactid" link-type="in" alias="i"><attribute name="name" /></link-entity></entity></fetch>`);
		expect(keys(fetch, [{ fullname: "F" }])).toEqual(["fullname"]);
	});

	it("keys Dataverse adds that the query didn't ask for (the primary key) aren't columns", () => {
		expect(keys(parse(ACCOUNT), [{ "@odata.etag": "x", accountid: "1", name: "A" }])).toEqual(["name", "primarycontactid", "address1_city"]);
	});

	it("without a query, every data key of every row", () => {
		expect(keys(null, [{ a: 1, "a@OData.Community.Display.V1.FormattedValue": "1" }, { b: 2 }])).toEqual(["a", "b"]);
	});
});

describe("resultColumns: matchfirstrowusingcrossapply (schema name keys)", () => {
	const XML = `<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="matchfirstrowusingcrossapply" alias="mc"><attribute name="fullname" /><attribute name="name" alias="n" /></link-entity></entity></fetch>`;

	it("uses the schema name from metadata", () => {
		const columns = resultColumns(parse(XML), [], meta({ "contact.fullname": { SchemaName: "FullName" } }));
		expect(columns.map((c) => c.key)).toEqual(["name", "FullName", "n"]);
		expect(columns[1].column?.key).toBe("mc.fullname");
	});

	it("without metadata, finds the key whose lowercase is the logical name", () => {
		expect(keys(parse(XML), [{ name: "A", FullName: "F" }])).toEqual(["name", "FullName", "n"]);
	});

	it("never takes a root attribute's key", () => {
		const fetch = parse(`<fetch><entity name="contact"><attribute name="fullname" /><link-entity name="contact" from="contactid" to="parentcustomerid" link-type="matchfirstrowusingcrossapply"><attribute name="fullname" /></link-entity></entity></fetch>`);
		expect(keys(fetch, [{ fullname: "Root", FullName: "Linked" }])).toEqual(["fullname", "FullName"]);
	});

	it("resolveColumn maps the schema name back, so sorting and filtering work", () => {
		const fetch = parse(XML);
		expect(resolveColumn(fetch, "FullName")?.attribute.name).toBe("fullname");
		expect(planSortByColumn(fetch, "FullName")).toEqual({ spec: { attribute: "fullname", entityname: "mc" } });
		// A root key never maps to the link-entity
		expect(resolveColumn(fetch, "name")?.isRoot).toBe(true);
	});
});

describe("resultColumns: all-attributes and paging", () => {
	it("<all-attributes /> takes the keys of every row, so a column null in row 1 is there", () => {
		const fetch = parse(`<fetch><entity name="account"><all-attributes /></entity></fetch>`);
		const rows = [{ accountid: "1", name: "A" }, { accountid: "2", name: "B", address1_city: "X" }];
		expect(keys(fetch, rows)).toEqual(["accountid", "name", "address1_city"]);
	});

	it("a root entity without attribute elements returns all its attributes", () => {
		const fetch = parse(`<fetch><entity name="account"><link-entity name="contact" from="contactid" to="primarycontactid" alias="c"><attribute name="fullname" /></link-entity></entity></fetch>`);
		expect(keys(fetch, [{ accountid: "1", name: "A", "c.fullname": "F" }])).toEqual(["c.fullname", "accountid", "name"]);
	});

	it("a link-entity's <all-attributes /> takes its prefixed keys only", () => {
		const fetch = parse(`<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer"><all-attributes /></link-entity></entity></fetch>`);
		const rows = [{ accountid: "1", name: "A", "contact1.fullname": "F" }, { accountid: "2", name: "B", "contact1.emailaddress1": "e" }];
		expect(keys(fetch, rows)).toEqual(["name", "contact1.fullname", "contact1.emailaddress1"]);
	});

	it("a cross-apply <all-attributes /> takes the schema name keys", () => {
		const fetch = parse(`<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="matchfirstrowusingcrossapply"><all-attributes /></link-entity></entity></fetch>`);
		expect(keys(fetch, [{ accountid: "1", name: "A", FullName: "F", ContactId: "c" }])).toEqual(["name", "FullName", "ContactId"]);
	});

	it("later pages add columns at the end and never move or drop earlier ones", () => {
		const fetch = parse(`<fetch><entity name="account"><attribute name="name" /><all-attributes /></entity></fetch>`);
		const page1 = [{ name: "A", accountid: "1" }];
		const page2 = [{ name: "B", accountid: "2", _primarycontactid_value: "c", revenue: 5 }];
		const before = keys(fetch, page1);
		const after = keys(fetch, [...page1, ...page2]);
		expect(after.slice(0, before.length)).toEqual(before);
		expect(after).toEqual(["name", "accountid", "_primarycontactid_value", "revenue"]);
	});

	it("rowKeys leaves annotations out and keeps first-seen order", () => {
		expect(rowKeys([{ b: 1, "b@x": 1 }, { a: 1, b: 2 }])).toEqual(["b", "a"]);
	});
});

describe("resultColumns: the tree changed after the run", () => {
	const ran = parse(`<fetch><entity name="account"><attribute name="name" /><attribute name="revenue" /></entity></fetch>`);
	const rows = [{ accountid: "1", name: "A" }];

	it("an attribute added since isn't an empty column until the query runs again", () => {
		const now = parse(`<fetch><entity name="account"><attribute name="name" /><attribute name="revenue" /><attribute name="telephone1" /></entity></fetch>`);
		expect(keys(now, rows, undefined, ran)).toEqual(["name", "revenue"]);
	});

	it("an attribute the rows already have shows straight away (the primary key)", () => {
		const now = parse(`<fetch><entity name="account"><attribute name="name" /><attribute name="accountid" /></entity></fetch>`);
		expect(keys(now, rows, undefined, ran)).toEqual(["name", "accountid"]);
	});

	it("a removed attribute goes", () => {
		const now = parse(`<fetch><entity name="account"><attribute name="revenue" /></entity></fetch>`);
		expect(keys(now, rows, undefined, ran)).toEqual(["revenue"]);
	});
});

describe("orderByLayout", () => {
	const layout = (...names: string[]): LayoutXmlConfig => ({ gridName: "resultset", columns: names.map((name) => ({ name, width: 100 })) });

	it("follows the layout; columns it doesn't have go last in their own order", () => {
		const fetch = parse(ACCOUNT);
		const columns = resultColumns(fetch, [{ name: "A", _primarycontactid_value: "c" }]);
		// The layout names the lookup by its attribute name
		expect(orderByLayout(columns, layout("primarycontactid", "name")).map((c) => c.key)).toEqual(["_primarycontactid_value", "name", "address1_city"]);
	});

	it("matches cross-apply columns by their query key, and all-attributes lookups by name", () => {
		const fetch = parse(`<fetch><entity name="account"><all-attributes /><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="matchfirstrowusingcrossapply" alias="mc"><attribute name="fullname" /></link-entity></entity></fetch>`);
		const columns = resultColumns(fetch, [{ name: "A", FullName: "F", _ownerid_value: "o" }]);
		expect(orderByLayout(columns, layout("ownerid", "mc.fullname")).map((c) => c.key)).toEqual(["_ownerid_value", "FullName", "name"]);
	});

	it("no layout keeps the order", () => {
		const columns = resultColumns(parse(ACCOUNT), []);
		expect(orderByLayout(columns, null)).toBe(columns);
	});
});
