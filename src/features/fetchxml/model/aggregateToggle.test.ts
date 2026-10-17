import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { planAggregateOff, hasAggregateOffChanges } from "./aggregateToggle";
import { findNodeContainer, collectAllLinkEntities, collectAllFilters } from "./treeUtils";

const tree = (xml: string) => parseFetchXml(xml).fetchNode!;

describe("planAggregateOff (#90b D4)", () => {
	const fetch = tree(
		`<fetch aggregate="true" aggregatelimit="500"><entity name="systemuser">` +
			`<attribute name="isdisabled" alias="d" groupby="true" />` +
			`<attribute name="systemuserid" alias="n" aggregate="count" />` +
			`<attribute name="createdon" alias="m" groupby="true" dategrouping="month" usertimezone="false" />` +
			`<order alias="n" descending="true" /><order alias="d" /><order alias="g" />` +
			`<link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="bu"><attribute name="name" alias="g" groupby="true" /></link-entity>` +
			`</entity></fetch>`,
	);
	const plan = planAggregateOff(fetch);

	it("clears aggregate-only settings on every attribute", () => {
		const attrIds = [...fetch.entity.attributes, ...fetch.entity.links[0].attributes].map((a) => a.id);
		expect(plan.updates.filter((u) => attrIds.includes(u.nodeId))).toHaveLength(4);
		expect(plan.removed.some((r) => r.message.includes('dategrouping="month"'))).toBe(true);
	});

	it("converts orders on Group by aliases to attribute orders", () => {
		const orderD = fetch.entity.orders[1];
		expect(plan.updates).toContainEqual({ nodeId: orderD.id, updates: { alias: undefined, attribute: "isdisabled" } });
	});

	it("removes orders on aggregate values and on aliases outside their scope", () => {
		const [orderN, , orderG] = fetch.entity.orders;
		expect(plan.removeNodeIds).toEqual([orderN.id, orderG.id]);
		expect(plan.removed.find((r) => r.message.startsWith("Sort by g"))?.message).toMatch(/isn't defined/);
	});

	it("drops aggregatelimit", () => {
		expect(plan.removed.some((r) => r.element === "fetch")).toBe(true);
	});

	it("has nothing to do for a plain query", () => {
		expect(hasAggregateOffChanges(planAggregateOff(tree(`<fetch><entity name="account"><attribute name="name" /></entity></fetch>`)))).toBe(false);
	});
});

describe("tree helpers", () => {
	const fetch = tree(
		`<fetch><entity name="account"><attribute name="name" />` +
			`<filter type="and"><condition attribute="statecode" operator="eq" value="0" />` +
			`<filter type="or"><condition attribute="name" operator="like" value="a%" /></filter>` +
			`<link-entity name="contact" from="parentcustomerid" to="accountid" link-type="any" alias="anyc"><filter><condition attribute="statecode" operator="eq" value="0" /></filter></link-entity>` +
			`</filter>` +
			`<link-entity name="contact" from="parentcustomerid" to="accountid" alias="c"><attribute name="fullname" /><order attribute="fullname" /></link-entity>` +
			`</entity></fetch>`,
	);

	it("finds the container and filter depth of a node", () => {
		const nested = fetch.entity.filters[0].subfilters[0];
		expect(findNodeContainer(fetch, nested.id)).toMatchObject({ isRoot: true, filterDepth: 2 });
		const linkOrder = fetch.entity.links[0].orders[0];
		expect(findNodeContainer(fetch, linkOrder.id)?.container.type).toBe("link-entity");
		expect(findNodeContainer(fetch, linkOrder.id)?.isRoot).toBe(false);
	});

	it("collects links and filters, including those inside filters", () => {
		expect(collectAllLinkEntities(fetch).map((l) => l.alias)).toEqual(["anyc", "c"]);
		expect(collectAllFilters(fetch)).toHaveLength(3);
	});
});

describe("findOrdersUsingAlias", () => {
	it("finds orders in the same scope, and root orders pointing at the link with entityname", async () => {
		const { findOrdersUsingAlias } = await import("./treeUtils");
		const f = tree(
			`<fetch aggregate="true"><entity name="systemuser"><attribute name="systemuserid" alias="n" aggregate="count" />` +
				`<order alias="n" /><order entityname="bu" alias="g" /><order alias="g" />` +
				`<link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="bu"><attribute name="name" alias="g" groupby="true" /><order alias="g" /></link-entity>` +
				`</entity></fetch>`,
		);
		const link = f.entity.links[0];
		expect(findOrdersUsingAlias(f, f.entity, "n").map((o) => o.id)).toEqual([f.entity.orders[0].id]);
		// root <order alias="g"> without entityname is NOT in the link's scope
		expect(findOrdersUsingAlias(f, link, "g").map((o) => o.id)).toEqual([link.orders[0].id, f.entity.orders[1].id]);
	});
});
