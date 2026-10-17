import { describe, expect, it } from "vitest";
import { addPagingToFetchXml, generateFetchXml } from "./fetchxml";
import type { EntityNode, FetchNode, LinkEntityNode } from "./nodes";

let id = 0;
const nid = () => `t_${++id}`;

function entity(name: string, overrides: Partial<EntityNode> = {}): EntityNode {
	return {
		id: nid(),
		type: "entity",
		name,
		attributes: [],
		orders: [],
		filters: [],
		links: [],
		...overrides,
	};
}

function fetch(root: EntityNode, options: FetchNode["options"] = {}): FetchNode {
	return { id: nid(), type: "fetch", entity: root, options };
}

function link(overrides: Partial<LinkEntityNode> & Pick<LinkEntityNode, "name" | "from" | "to">): LinkEntityNode {
	return {
		id: nid(),
		type: "link-entity",
		linkType: "inner",
		attributes: [],
		orders: [],
		filters: [],
		links: [],
		...overrides,
	};
}

describe("generateFetchXml", () => {
	it("serializes fetch options", () => {
		const xml = generateFetchXml(
			fetch(entity("account"), { top: 50, distinct: true, noLock: true, returnTotalRecordCount: true }),
		);
		expect(xml.split("\n")[0]).toBe(
			'<fetch distinct="true" top="50" returntotalrecordcount="true" no-lock="true">',
		);
	});

	it("serializes attributes, orders, filters and conditions", () => {
		const xml = generateFetchXml(
			fetch(
				entity("account", {
					attributes: [
						{ id: nid(), type: "attribute", name: "name" },
						{ id: nid(), type: "attribute", name: "revenue", alias: "rev" },
					],
					orders: [{ id: nid(), type: "order", attribute: "name", descending: true }],
					filters: [
						{
							id: nid(),
							type: "filter",
							conjunction: "and",
							conditions: [
								{ id: nid(), type: "condition", attribute: "statecode", operator: "eq", value: 0 },
								{ id: nid(), type: "condition", attribute: "parentaccountid", operator: "null" },
							],
							subfilters: [
								{
									id: nid(),
									type: "filter",
									conjunction: "or",
									conditions: [
										{ id: nid(), type: "condition", attribute: "name", operator: "like", value: "a%" },
									],
									subfilters: [],
									links: [],
								},
							],
							links: [],
						},
					],
				}),
			),
		);

		expect(xml).toBe(
			[
				"<fetch>",
				'  <entity name="account">',
				'    <attribute name="name" />',
				'    <attribute name="revenue" alias="rev" />',
				'    <order attribute="name" descending="true" />',
				'    <filter type="and">',
				'      <condition attribute="statecode" operator="eq" value="0" />',
				'      <condition attribute="parentaccountid" operator="null" />',
				'      <filter type="or">',
				'        <condition attribute="name" operator="like" value="a%" />',
				"      </filter>",
				"    </filter>",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
	});

	it("serializes multi-value conditions as <value> children", () => {
		const xml = generateFetchXml(
			fetch(
				entity("account", {
					filters: [
						{
							id: nid(),
							type: "filter",
							conjunction: "and",
							conditions: [
								{ id: nid(), type: "condition", attribute: "industrycode", operator: "in", value: [1, 2, 3] },
							],
							subfilters: [],
							links: [],
						},
					],
				}),
			),
		);
		expect(xml).toContain('<condition attribute="industrycode" operator="in">');
		expect(xml).toContain("<value>1</value>");
		expect(xml).toContain("<value>3</value>");
	});

	it("serializes link-entities with non-default link type and alias", () => {
		const xml = generateFetchXml(
			fetch(
				entity("contact", {
					links: [
						link({
							name: "account",
							from: "accountid",
							to: "parentcustomerid",
							alias: "acc",
							linkType: "outer",
							attributes: [{ id: nid(), type: "attribute", name: "name" }],
						}),
					],
				}),
			),
		);
		expect(xml).toContain(
			'<link-entity name="account" from="accountid" to="parentcustomerid" alias="acc" link-type="outer">',
		);
		expect(xml).toContain('      <attribute name="name" />');
	});

	it("omits link-type for inner joins", () => {
		const xml = generateFetchXml(
			fetch(entity("contact", { links: [link({ name: "account", from: "accountid", to: "parentcustomerid" })] })),
		);
		expect(xml).toContain('<link-entity name="account" from="accountid" to="parentcustomerid">');
	});

	it("escapes XML special characters in condition values", () => {
		const xml = generateFetchXml(
			fetch(
				entity("account", {
					filters: [
						{
							id: nid(),
							type: "filter",
							conjunction: "and",
							conditions: [
								{ id: nid(), type: "condition", attribute: "name", operator: "eq", value: 'A & B <"C">' },
							],
							subfilters: [],
							links: [],
						},
					],
				}),
			),
		);
		expect(xml).toContain('value="A &amp; B &lt;&quot;C&quot;&gt;"');
	});

	it("avoids scientific notation for small numbers", () => {
		const xml = generateFetchXml(
			fetch(
				entity("account", {
					filters: [
						{
							id: nid(),
							type: "filter",
							conjunction: "and",
							conditions: [
								{ id: nid(), type: "condition", attribute: "revenue", operator: "gt", value: 0.0000000000012 },
							],
							subfilters: [],
							links: [],
						},
					],
				}),
			),
		);
		expect(xml).not.toMatch(/e-/);
		expect(xml).toContain('value="0.0000000000012"');
	});
});

describe("generateFetchXml never adds attributes (#54, #59)", () => {
	const attributeNames = (xml: string) => [...xml.matchAll(/<attribute name="([^"]+)"/g)].map((m) => m[1]);

	it.each([
		["plain", {}],
		["aggregate", { aggregate: true }],
		["distinct", { distinct: true }],
	] as const)("%s query contains only the tree's attributes", (_label, options) => {
		const xml = generateFetchXml(
			fetch(
				entity("account", {
					attributes: [{ id: nid(), type: "attribute", name: "name" }],
					links: [
						link({
							name: "contact",
							from: "parentcustomerid",
							to: "accountid",
							alias: "c",
							attributes: [{ id: nid(), type: "attribute", name: "fullname" }],
						}),
					],
				}),
				options,
			),
		);
		expect(attributeNames(xml)).toEqual(["name", "fullname"]);
	});
});

describe("addPagingToFetchXml", () => {
	const base = '<fetch count="50">\n  <entity name="account" />\n</fetch>';

	it("adds page and count", () => {
		expect(addPagingToFetchXml(base, 2, undefined, 50).split("\n")[0]).toMatch(
			/^<fetch\s+page="2" count="50">$/,
		);
	});

	it("replaces existing page, count and paging-cookie attributes", () => {
		const first = addPagingToFetchXml(base, 2, '<cookie page="1"><accountid last="{A}" /></cookie>', 50);
		const second = addPagingToFetchXml(first, 3, '<cookie page="2"><accountid last="{B}" /></cookie>', 50);
		const fetchTag = second.split("\n")[0];
		expect(fetchTag.match(/page="/g)).toHaveLength(1);
		expect(fetchTag.match(/count="/g)).toHaveLength(1);
		expect(fetchTag.match(/paging-cookie="/g)).toHaveLength(1);
		expect(fetchTag).toContain('page="3"');
		expect(fetchTag).toContain("{B}");
		expect(fetchTag).not.toContain("{A}");
	});

	it("XML-escapes the paging cookie", () => {
		const xml = addPagingToFetchXml(base, 2, '<cookie page="1" />');
		expect(xml).toContain('paging-cookie="&lt;cookie page=&quot;1&quot; /&gt;"');
	});

	it("returns input unchanged when there is no fetch element", () => {
		expect(addPagingToFetchXml("<entity />", 2)).toBe("<entity />");
	});
});
