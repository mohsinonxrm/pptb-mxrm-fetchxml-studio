/**
 * Result columns ↔ tree (#114). Column keys mirror what Dataverse returned live (2026-10-15).
 */
import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import {
	collectQueryColumns,
	planRemoveColumns,
	planSortByColumn,
	resolveColumn,
} from "./queryColumns";

const parse = (xml: string) => parseFetchXml(xml).fetchNode!;
const OWN = (alias = "", inner = `<attribute name="fullname" />`) =>
	`<link-entity name="systemuser" from="systemuserid" to="owninguser" link-type="outer"${alias}>${inner}</link-entity>`;
const BU = `<link-entity name="businessunit" from="businessunitid" to="owningbusinessunit" link-type="outer"><attribute name="name" /></link-entity>`;

describe("collectQueryColumns: keys as Dataverse returns them", () => {
	it.each([
		[
			"two unaliased link-entities",
			`<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer"><attribute name="fullname" /></link-entity>${OWN()}</entity></fetch>`,
			["name", "contact1.fullname", "systemuser2.fullname"],
		],
		[
			"aliased first, unaliased second",
			`<fetch><entity name="account"><attribute name="name" />${OWN(` alias="o"`)}${BU}</entity></fetch>`,
			["name", "o.fullname", "businessunit2.name"],
		],
		[
			"link-entity inside a filter isn't counted",
			`<fetch><entity name="account"><attribute name="name" /><filter><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="any" /></filter>${OWN()}</entity></fetch>`,
			["name", "systemuser1.fullname"],
		],
		[
			"nested unaliased under unaliased",
			`<fetch><entity name="account"><attribute name="name" />${OWN("", `<attribute name="fullname" />${BU.replace('to="owningbusinessunit"', 'to="businessunitid"')}`)}</entity></fetch>`,
			["name", "systemuser1.fullname", "businessunit2.name"],
		],
		[
			"attribute alias wins",
			`<fetch><entity name="account"><attribute name="name" alias="n" />${OWN(` alias="o"`, `<attribute name="fullname" alias="owner" />`)}</entity></fetch>`,
			["n", "owner"],
		],
	])("%s", (_label, xml, keys) => {
		expect(collectQueryColumns(parse(xml)).map((c) => c.key)).toEqual(keys);
	});

	it("resolves the lookup form _name_value to the root attribute", () => {
		const fetch = parse(
			`<fetch><entity name="account"><attribute name="primarycontactid" /></entity></fetch>`,
		);
		expect(resolveColumn(fetch, "_primarycontactid_value")?.attribute.name).toBe(
			"primarycontactid",
		);
	});
});

describe("planRemoveColumns", () => {
	it("removes aliased root attributes and link-entity attributes (both were ignored before)", () => {
		const fetch = parse(
			`<fetch><entity name="account"><attribute name="name" alias="n" /><attribute name="accountnumber" />${OWN(` alias="o"`, `<attribute name="fullname" /><attribute name="domainname" />`)}</entity></fetch>`,
		);
		const plan = planRemoveColumns(fetch, ["n", "o.fullname"]);
		expect(plan.removeNodeIds).toEqual([
			fetch.entity.attributes[0].id,
			fetch.entity.links[0].attributes[0].id,
		]);
		expect(plan.removedLinks).toEqual([]);
	});

	it("removes sorts that use a removed attribute's alias (aggregate)", () => {
		const fetch = parse(
			`<fetch aggregate="true"><entity name="account"><attribute name="industrycode" alias="ind" groupby="true" /><attribute name="accountid" alias="n" aggregate="count" /><order alias="n" descending="true" /></entity></fetch>`,
		);
		const plan = planRemoveColumns(fetch, ["n"]);
		expect(plan.removeNodeIds).toEqual(
			expect.arrayContaining([fetch.entity.attributes[1].id, fetch.entity.orders[0].id]),
		);
	});

	it("removes an outer link-entity left with nothing in it (a 1:N join would keep multiplying rows)", () => {
		const fetch = parse(
			`<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="outer" alias="c"><attribute name="fullname" /></link-entity></entity></fetch>`,
		);
		const plan = planRemoveColumns(fetch, ["c.fullname"]);
		expect(plan.removeNodeIds).toEqual([fetch.entity.links[0].id]);
		expect(plan.removedLinks.map((l) => l.alias)).toEqual(["c"]);
	});

	it("keeps inner joins, and link-entities that still filter or have other content", () => {
		const inner = parse(
			`<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="contactid" to="primarycontactid" alias="c"><attribute name="fullname" /></link-entity></entity></fetch>`,
		);
		expect(planRemoveColumns(inner, ["c.fullname"]).removedLinks).toEqual([]);
		const filtered = parse(
			`<fetch><entity name="account"><attribute name="name" /><link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="c"><attribute name="fullname" /><filter><condition attribute="statecode" operator="eq" value="0" /></filter></link-entity></entity></fetch>`,
		);
		expect(planRemoveColumns(filtered, ["c.fullname"]).removedLinks).toEqual([]);
	});

	it("removes the emptied child link-entity and keeps its parent (which may be there on purpose)", () => {
		const fetch = parse(
			`<fetch><entity name="account"><attribute name="name" />${OWN(` alias="o"`, `<link-entity name="businessunit" from="businessunitid" to="businessunitid" link-type="outer" alias="b"><attribute name="name" /></link-entity>`)}</entity></fetch>`,
		);
		const plan = planRemoveColumns(fetch, ["b.name"]);
		expect(plan.removedLinks.map((l) => l.alias)).toEqual(["b"]);
		// The parent had no attribute of its own removed, so it stays (it may have been added on purpose)
		expect(plan.removeNodeIds).toEqual([fetch.entity.links[0].links[0].id]);
	});
});

describe("planRemoveColumns: root entityname references to the link-entity (#357)", () => {
	const ACCOUNT = (inner: string) =>
		parse(`<fetch><entity name="account"><attribute name="name" />${inner}</entity></fetch>`);
	const ids = (...nodes: { id: string }[]) => nodes.map((n) => n.id).sort();

	it.each([
		["by name, link-entity without an alias", "", "systemuser", "systemuser1.fullname"],
		["by alias", ` alias="u"`, "u", "u.fullname"],
		["by name, link-entity with an alias", ` alias="u"`, "systemuser", "u.fullname"],
	])(
		"a root order on the removed column goes with the link-entity (%s)",
		(_label, alias, entityname, key) => {
			const fetch = ACCOUNT(
				`${OWN(alias)}<order attribute="fullname" entityname="${entityname}" />`,
			);
			const plan = planRemoveColumns(fetch, [key]);
			expect(plan.removedLinks).toEqual([fetch.entity.links[0]]);
			expect([...plan.removeNodeIds].sort()).toEqual(
				ids(fetch.entity.links[0], fetch.entity.orders[0]),
			);
		},
	);

	it("an order inside the link-entity on the removed column goes with it too", () => {
		const fetch = ACCOUNT(
			OWN(` alias="u"`, `<attribute name="fullname" /><order attribute="fullname" />`),
		);
		const plan = planRemoveColumns(fetch, ["u.fullname"]);
		expect(plan.removedLinks).toEqual([fetch.entity.links[0]]);
		// The order is inside the removed link-entity, so it isn't listed on its own
		expect(plan.removeNodeIds).toEqual([fetch.entity.links[0].id]);
	});

	it("a root order on another attribute keeps the link-entity, and the order (owner decision 2)", () => {
		const fetch = ACCOUNT(`${OWN(` alias="u"`)}<order attribute="createdon" entityname="u" />`);
		const plan = planRemoveColumns(fetch, ["u.fullname"]);
		expect(plan.removedLinks).toEqual([]);
		expect(plan.removeNodeIds).toEqual([fetch.entity.links[0].attributes[0].id]);
	});

	it("an order inside the link-entity on another attribute keeps it", () => {
		const fetch = ACCOUNT(
			OWN(` alias="u"`, `<attribute name="fullname" /><order attribute="createdon" />`),
		);
		expect(planRemoveColumns(fetch, ["u.fullname"]).removedLinks).toEqual([]);
	});

	it.each([
		[
			"in the root filter",
			`<filter><condition entityname="u" attribute="fullname" operator="like" value="%a%" /></filter>`,
		],
		[
			"in a nested filter",
			`<filter type="or"><condition attribute="name" operator="null" /><filter><condition entityname="u" attribute="fullname" operator="like" value="%a%" /></filter></filter>`,
		],
		[
			"by the link-entity's name",
			`<filter><condition entityname="systemuser" attribute="fullname" operator="not-null" /></filter>`,
		],
	])(
		"a root condition on the link-entity keeps it, so the rows stay filtered (%s)",
		(_label, filter) => {
			const fetch = ACCOUNT(
				`${OWN(` alias="u"`)}${filter}<order attribute="fullname" entityname="u" />`,
			);
			const plan = planRemoveColumns(fetch, ["u.fullname"]);
			expect(plan.removedLinks).toEqual([]);
			// Only the attribute goes: the link-entity stays, so its order is still valid
			expect(plan.removeNodeIds).toEqual([fetch.entity.links[0].attributes[0].id]);
		},
	);

	it("resolves entityname to an alias before a name, as the order rules do", () => {
		// The order's entityname is the contact link-entity's alias, not the removed systemuser link-entity
		const fetch = ACCOUNT(
			`<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="systemuser"><attribute name="fullname" /></link-entity>${OWN(` alias="u"`)}<order attribute="fullname" entityname="systemuser" />`,
		);
		const plan = planRemoveColumns(fetch, ["u.fullname"]);
		expect(plan.removedLinks).toEqual([fetch.entity.links[1]]);
		expect(plan.removeNodeIds).toEqual([fetch.entity.links[1].id]);
	});

	it("a root order on a root attribute isn't touched", () => {
		const fetch = ACCOUNT(`${OWN(` alias="u"`)}<order attribute="fullname" />`);
		const plan = planRemoveColumns(fetch, ["u.fullname"]);
		expect(plan.removeNodeIds).toEqual([fetch.entity.links[0].id]);
	});

	it("self-join: entityname with the root's name is the root entity, so its order stays (verified live, #401)", () => {
		const fetch = ACCOUNT(
			`<link-entity name="account" from="accountid" to="parentaccountid" link-type="outer"><attribute name="name" /></link-entity><order attribute="name" entityname="account" />`,
		);
		const plan = planRemoveColumns(fetch, ["account1.name"]);
		expect(plan.removedLinks).toEqual([fetch.entity.links[0]]);
		expect(plan.removeNodeIds).toEqual([fetch.entity.links[0].id]);
	});
});

describe("planRemoveColumns: conditions that compare to the link-entity through valueof (#357)", () => {
	const C = (inner: string) =>
		`<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="c">${inner}</link-entity>`;
	const ACCOUNT = (inner: string) =>
		parse(`<fetch><entity name="account"><attribute name="name" />${inner}</entity></fetch>`);
	const COMPARE = `<filter><condition attribute="fullname" operator="eq" valueof="u.fullname" /></filter>`;

	it.each([
		[
			"in the root filter",
			`${OWN(` alias="u"`)}<filter><condition attribute="name" operator="eq" valueof="u.fullname" /></filter>`,
		],
		[
			"in a nested root filter",
			`${OWN(` alias="u"`)}<filter type="or"><condition attribute="name" operator="null" /><filter><condition attribute="name" operator="ne" valueof="u.fullname" /></filter></filter>`,
		],
		[
			"in a later link-entity's filter",
			`${OWN(` alias="u"`)}${C(`<attribute name="fullname" />${COMPARE}`)}`,
		],
		[
			"in a link-entity inside a filter (any)",
			`${OWN(` alias="u"`)}<filter><link-entity name="contact" from="parentcustomerid" to="accountid" link-type="any">${COMPARE}</link-entity></filter>`,
		],
	])("keeps the link-entity, so the comparison still has its attribute (%s)", (_label, inner) => {
		const fetch = ACCOUNT(inner);
		const link = fetch.entity.links.find((l) => l.alias === "u")!;
		const plan = planRemoveColumns(fetch, ["u.fullname"]);
		expect(plan.removedLinks).toEqual([]);
		expect(plan.removeNodeIds).toEqual([link.attributes[0].id]);
	});

	it("a valueof on the same entity, or on another alias, doesn't keep it", () => {
		const fetch = ACCOUNT(
			`${OWN(` alias="u"`)}${C(`<attribute name="fullname" />`)}<filter><condition attribute="name" operator="eq" valueof="accountnumber" /><condition attribute="name" operator="eq" valueof="c.fullname" /></filter>`,
		);
		expect(planRemoveColumns(fetch, ["u.fullname"]).removedLinks).toEqual([
			fetch.entity.links[0],
		]);
	});

	it("matches the alias only, not the link-entity's name", () => {
		const fetch = ACCOUNT(
			`${OWN(` alias="u"`)}<filter><condition attribute="name" operator="eq" valueof="systemuser.fullname" /></filter>`,
		);
		expect(planRemoveColumns(fetch, ["u.fullname"]).removedLinks).toEqual([
			fetch.entity.links[0],
		]);
	});
});

describe("planSortByColumn", () => {
	const fetch = parse(
		`<fetch><entity name="account"><attribute name="name" alias="n" />${OWN(` alias="o"`, `<attribute name="fullname" alias="owner" />`)}<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer"><attribute name="fullname" /></link-entity></entity></fetch>`,
	);
	it("root attribute with an alias sorts by attribute name", () => {
		expect(planSortByColumn(fetch, "n")).toEqual({ spec: { attribute: "name" } });
	});
	it("aliased link-entity attribute sorts from the root with entityname (verified; alias fails 0x80041103)", () => {
		expect(planSortByColumn(fetch, "owner")).toEqual({
			spec: { attribute: "fullname", entityname: "o" },
		});
	});
	it("unaliased link-entity with a unique name uses its name", () => {
		expect(planSortByColumn(fetch, "contact2.fullname")).toEqual({
			spec: { attribute: "fullname", entityname: "contact" },
		});
	});
	it("unaliased link-entity with a repeated name sorts inside the link-entity", () => {
		const twice = parse(
			`<fetch><entity name="account"><attribute name="name" />${OWN()}<link-entity name="systemuser" from="systemuserid" to="createdby" link-type="outer"><attribute name="fullname" /></link-entity></entity></fetch>`,
		);
		const plan = planSortByColumn(twice, "systemuser2.fullname");
		expect(plan).toMatchObject({
			spec: { attribute: "fullname" },
			insideLink: { id: twice.entity.links[1].id },
		});
	});
	it("aggregate queries sort by alias", () => {
		const agg = parse(
			`<fetch aggregate="true"><entity name="account"><attribute name="industrycode" alias="ind" groupby="true" /><attribute name="accountid" alias="n" aggregate="count" /></entity></fetch>`,
		);
		expect(planSortByColumn(agg, "n")).toEqual({ spec: { alias: "n" } });
	});
	// #92: aggregate queries sort by alias, also for link-entity columns. An aggregate order alias is
	// resolved in the order's own scope, or in the entityname link-entity (verified, fetchxmlRules.ts).
	describe("aggregate link-entity columns (#92)", () => {
		const agg = parse(
			`<fetch aggregate="true"><entity name="opportunity"><attribute name="opportunityid" alias="n" aggregate="count" /><link-entity name="account" from="accountid" to="parentaccountid" alias="a"><attribute name="name" alias="acct" groupby="true" /></link-entity><link-entity name="systemuser" from="systemuserid" to="owninguser"><attribute name="fullname" alias="owner" groupby="true" /></link-entity></entity></fetch>`,
		);
		it("aliased link-entity: alias with entityname, never attribute", () => {
			expect(planSortByColumn(agg, "acct")).toEqual({
				spec: { alias: "acct", entityname: "a" },
			});
		});
		it("unaliased link-entity with a unique name: alias with its name", () => {
			expect(planSortByColumn(agg, "owner")).toEqual({
				spec: { alias: "owner", entityname: "systemuser" },
			});
		});
		it("unaliased link-entity with a repeated name: alias inside the link-entity", () => {
			const twice = parse(
				`<fetch aggregate="true"><entity name="account"><attribute name="accountid" alias="n" aggregate="count" /><link-entity name="systemuser" from="systemuserid" to="owninguser"><attribute name="fullname" alias="o" groupby="true" /></link-entity><link-entity name="systemuser" from="systemuserid" to="createdby"><attribute name="fullname" alias="c" groupby="true" /></link-entity></entity></fetch>`,
			);
			expect(planSortByColumn(twice, "c")).toMatchObject({
				spec: { alias: "c" },
				insideLink: { id: twice.entity.links[1].id },
			});
		});
		it("an attribute without an alias isn't sortable in an aggregate query", () => {
			const noAlias = parse(
				`<fetch aggregate="true"><entity name="account"><attribute name="industrycode" groupby="true" /><attribute name="accountid" alias="n" aggregate="count" /></entity></fetch>`,
			);
			expect(planSortByColumn(noAlias, "industrycode")).toHaveProperty("reason");
		});
	});
	it("root lookups sort by attribute name from their _name_value column (#91)", () => {
		const lookup = parse(
			`<fetch><entity name="account"><attribute name="primarycontactid" /></entity></fetch>`,
		);
		expect(planSortByColumn(lookup, "_primarycontactid_value")).toEqual({
			spec: { attribute: "primarycontactid" },
		});
	});
	it("multi-select option sets aren't sortable (0x80050224)", () => {
		const ms = parse(
			`<fetch><entity name="appaction"><attribute name="clienttype" /></entity></fetch>`,
		);
		expect(
			planSortByColumn(ms, "clienttype", { typeName: "MultiSelectPicklistType" }),
		).toHaveProperty("reason");
	});
});
