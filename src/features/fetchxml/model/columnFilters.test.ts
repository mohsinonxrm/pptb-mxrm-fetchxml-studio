import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { generateFetchXml } from "./fetchxml";
import {
	addColumnFilter,
	clearColumnFilters,
	columnFilterConditions,
	columnFilterTarget,
	relocateColumnFilterGroup,
	relocateColumnFilterWrapper,
	removeColumnFilter,
} from "./columnFilters";

const parse = (xml: string) => parseFetchXml(xml).fetchNode!;
let n = 0;
const id = () => `cf_${++n}`;
/** The generated FetchXML without whitespace between elements */
const xml = (fetch: Parameters<typeof generateFetchXml>[0]) => generateFetchXml(fetch).replace(/>\s+</g, "><");

const CITY = { attribute: "address1_city", operator: "eq" as const, value: "Redmond" };
const NAME = { attribute: "name", operator: "like" as const, value: "%Contoso%" };

describe("columnFilterTarget", () => {
	const fetch = parse(`<fetch><entity name="account">
		<attribute name="name" alias="n" />
		<attribute name="primarycontactid" />
		<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="pc"><attribute name="fullname" /></link-entity>
		<link-entity name="systemuser" from="systemuserid" to="owninguser"><attribute name="fullname" /></link-entity>
	</entity></fetch>`);

	it("targets root attributes by logical name, also when aliased or a lookup", () => {
		expect(columnFilterTarget(fetch, "n")).toEqual({ attribute: "name", entityName: "account" });
		expect(columnFilterTarget(fetch, "_primarycontactid_value")).toEqual({ attribute: "primarycontactid", entityName: "account" });
	});

	it("targets link-entity attributes with entityname: the alias, or the name without one", () => {
		expect(columnFilterTarget(fetch, "pc.fullname")).toEqual({ attribute: "fullname", entityname: "pc", entityName: "contact" });
		expect(columnFilterTarget(fetch, "systemuser2.fullname")).toEqual({
			attribute: "fullname",
			entityname: "systemuser",
			entityName: "systemuser",
		});
	});

	it("refuses unknown columns and aggregate queries", () => {
		expect(columnFilterTarget(fetch, "nope")).toHaveProperty("reason");
		const aggregate = parse(`<fetch aggregate="true"><entity name="account"><attribute name="name" alias="c" aggregate="count" /></entity></fetch>`);
		expect(columnFilterTarget(aggregate, "c")).toHaveProperty("reason");
	});
});

describe("addColumnFilter", () => {
	it("makes the group the root filter when there is none", () => {
		const fetch = parse(`<fetch><entity name="account"><attribute name="name" /></entity></fetch>`);
		const { fetch: next, groupId } = addColumnFilter(fetch, null, CITY, id);
		expect(xml(next)).toContain(`<filter type="and"><condition attribute="address1_city" operator="eq" value="Redmond" /></filter>`);
		expect(next.entity.filters).toHaveLength(1);
		expect(next.entity.filters[0].id).toBe(groupId);
	});

	it("nests the group in an 'and' root filter, leaving its conditions alone", () => {
		const fetch = parse(`<fetch><entity name="account"><filter><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>`);
		const { fetch: next } = addColumnFilter(fetch, null, CITY, id);
		expect(xml(next)).toContain(
			`<filter type="and"><condition attribute="statecode" operator="eq" value="0" /><filter type="and"><condition attribute="address1_city" operator="eq" value="Redmond" /></filter></filter>`,
		);
		expect(next.entity.filters).toHaveLength(1);
	});

	it("wraps an 'or' root filter so the column filter still narrows the results", () => {
		const fetch = parse(`<fetch><entity name="account"><filter type="or"><condition attribute="statecode" operator="eq" value="0" /><condition attribute="statecode" operator="eq" value="1" /></filter></entity></fetch>`);
		const { fetch: next } = addColumnFilter(fetch, null, CITY, id);
		expect(xml(next)).toContain(
			`<filter type="and"><filter type="or"><condition attribute="statecode" operator="eq" value="0" /><condition attribute="statecode" operator="eq" value="1" /></filter><filter type="and"><condition attribute="address1_city" operator="eq" value="Redmond" /></filter></filter>`,
		);
	});

	it("adds later filters to the same group", () => {
		const fetch = parse(`<fetch><entity name="account"><filter><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>`);
		const first = addColumnFilter(fetch, null, CITY, id);
		const second = addColumnFilter(first.fetch, first.groupId, NAME, id);
		expect(second.groupId).toBe(first.groupId);
		expect(columnFilterConditions(second.fetch, second.groupId).map((c) => c.attribute)).toEqual(["address1_city", "name"]);
	});

	it("writes entityname for link-entity columns", () => {
		const fetch = parse(`<fetch><entity name="account"><link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="pc" /></entity></fetch>`);
		const { fetch: next } = addColumnFilter(fetch, null, { attribute: "fullname", entityname: "pc", operator: "null" }, id);
		expect(xml(next)).toContain(`<condition attribute="fullname" operator="null" entityname="pc" />`);
	});

	it("starts a new group when the old one was removed in the tree", () => {
		const fetch = parse(`<fetch><entity name="account" /></fetch>`);
		const { fetch: next } = addColumnFilter(fetch, "gone", CITY, id);
		expect(columnFilterConditions(next, "gone")).toEqual([]);
		expect(next.entity.filters[0].conditions).toHaveLength(1);
	});
});

describe("removeColumnFilter / clearColumnFilters", () => {
	const base = parse(`<fetch><entity name="account"><filter><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>`);

	it("removes one condition and keeps the group", () => {
		const a = addColumnFilter(base, null, CITY, id);
		const b = addColumnFilter(a.fetch, a.groupId, NAME, id);
		const removed = removeColumnFilter(b.fetch, b.groupId, a.conditionId);
		expect(removed.groupId).toBe(a.groupId);
		expect(columnFilterConditions(removed.fetch, removed.groupId).map((c) => c.attribute)).toEqual(["name"]);
	});

	it("removes the empty group, restoring the original query", () => {
		const a = addColumnFilter(base, null, CITY, id);
		const removed = removeColumnFilter(a.fetch, a.groupId, a.conditionId);
		expect(removed.groupId).toBeNull();
		expect(xml(removed.fetch)).toBe(xml(base));
	});

	it("removes a group that was the root filter", () => {
		const plain = parse(`<fetch><entity name="account" /></fetch>`);
		const a = addColumnFilter(plain, null, CITY, id);
		expect(removeColumnFilter(a.fetch, a.groupId, a.conditionId).fetch.entity.filters).toEqual([]);
	});

	it("unwraps an 'or' root filter again when the last column filter goes", () => {
		const or = parse(`<fetch><entity name="account"><filter type="or"><condition attribute="statecode" operator="eq" value="0" /><condition attribute="statecode" operator="eq" value="1" /></filter></entity></fetch>`);
		const a = addColumnFilter(or, null, CITY, id);
		expect(a.wrapperId).toBe(a.fetch.entity.filters[0].id);
		expect(xml(removeColumnFilter(a.fetch, a.groupId, a.conditionId, a.wrapperId).fetch)).toBe(xml(or));
	});

	it("keeps an 'and' wrapper the user wrote, even when it has the same shape", () => {
		const authored = parse(`<fetch><entity name="account"><filter type="and"><filter type="or"><condition attribute="statecode" operator="eq" value="0" /><condition attribute="statecode" operator="eq" value="1" /></filter></filter></entity></fetch>`);
		const a = addColumnFilter(authored, null, CITY, id);
		expect(a.wrapperId).toBeUndefined();
		expect(xml(removeColumnFilter(a.fetch, a.groupId, a.conditionId, null).fetch)).toBe(xml(authored));
	});

	it("keeps the wrapper when something was added to it since", () => {
		const or = parse(`<fetch><entity name="account"><filter type="or"><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>`);
		const a = addColumnFilter(or, null, CITY, id);
		const root = a.fetch.entity.filters[0];
		const edited = {
			...a.fetch,
			entity: { ...a.fetch.entity, filters: [{ ...root, conditions: [{ id: "u1", type: "condition" as const, attribute: "name", operator: "not-null" as const }] }] },
		};
		const removed = removeColumnFilter(edited, a.groupId, a.conditionId, a.wrapperId).fetch;
		expect(removed.entity.filters[0].id).toBe(a.wrapperId);
	});

	it("clears every column filter", () => {
		const a = addColumnFilter(base, null, CITY, id);
		const b = addColumnFilter(a.fetch, a.groupId, NAME, id);
		expect(xml(clearColumnFilters(b.fetch, b.groupId))).toBe(xml(base));
	});
});

describe("relocateColumnFilterGroup", () => {
	it("finds the wrapper again after the query is parsed again", () => {
		const or = parse(`<fetch><entity name="account"><filter type="or"><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>`);
		const a = addColumnFilter(or, null, CITY, id);
		const reparsed = parse(generateFetchXml(a.fetch));
		const groupId = relocateColumnFilterGroup(a.fetch, reparsed, a.groupId);
		const wrapperId = relocateColumnFilterWrapper(a.fetch, reparsed, a.wrapperId!, groupId);
		expect(wrapperId).toBe(reparsed.entity.filters[0].id);
		expect(xml(removeColumnFilter(reparsed, groupId, columnFilterConditions(reparsed, groupId)[0].id, wrapperId).fetch)).toBe(xml(or));
	});

	it("finds the group again after the query is parsed again", () => {
		const base = parse(`<fetch><entity name="account"><filter><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>`);
		const a = addColumnFilter(base, null, CITY, id);
		const reparsed = parse(generateFetchXml(a.fetch));
		const relocated = relocateColumnFilterGroup(a.fetch, reparsed, a.groupId);
		expect(relocated).toBe(reparsed.entity.filters[0].subfilters[0].id);
		expect(columnFilterConditions(reparsed, relocated)).toHaveLength(1);
	});

	it("lets go when the group's conditions changed", () => {
		const base = parse(`<fetch><entity name="account" /></fetch>`);
		const a = addColumnFilter(base, null, CITY, id);
		const other = parse(`<fetch><entity name="account"><filter><condition attribute="name" operator="eq" value="x" /></filter></entity></fetch>`);
		expect(relocateColumnFilterGroup(a.fetch, other, a.groupId)).toBeNull();
	});
});
