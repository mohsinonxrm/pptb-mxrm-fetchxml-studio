/**
 * Building and editing queries from the tree (#345 review): queries built from an empty tree with
 * the same builder actions the tree's Add menus and the property editors use (BuilderProvider +
 * useBuilder), then checked as FetchXML, as query-rule findings, and through a FetchXML → tree →
 * FetchXML round trip.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { BuilderProvider } from "../state/builderStore";
import { useBuilder, type BuilderContextValue } from "../state/builderContext";
import { generateFetchXml } from "../model/fetchxml";
import { getExecutionErrors, parseFetchXml } from "../model/fetchxmlParser";
import { validateQueryRules } from "../model/fetchxmlRules";
import {
	collectAllLinkEntities,
	findParentEntityForNode,
	findOrdersUsingAlias,
} from "../model/treeUtils";
import type { FetchNode, LinkEntityNode, NodeId } from "../model/nodes";
import {
	compactFetchXml,
	generateFxsSyncCode,
	generatePowerAutomateSpec,
} from "../engine/fetchxmlCodeGenerators";
import { cleanup, render } from "../../../shared/testing/ui";

afterEach(() => cleanup());

// ---------------------------------------------------------------------------------------------
// Harness: the real BuilderProvider, driven through its context like the tree and editors do
// ---------------------------------------------------------------------------------------------

let current: BuilderContextValue | null = null;
function Probe() {
	current = useBuilder();
	return null;
}

interface Builder {
	/** The latest context value (it changes with every state change) */
	readonly b: BuilderContextValue;
	readonly query: FetchNode;
	/** Runs builder actions inside act; returns the selected node id afterwards (the added node) */
	run: (fn: (b: BuilderContextValue) => void) => Promise<NodeId>;
	xml: () => string;
	errors: () => { nodeId: string; attribute?: string; message: string }[];
}

async function startBuilder(rootEntity = "account"): Promise<Builder> {
	current = null;
	await render(createElement(BuilderProvider, null, createElement(Probe)));
	const builder: Builder = {
		get b() {
			if (!current) throw new Error("builder not rendered");
			return current;
		},
		get query() {
			const q = builder.b.fetchQuery;
			if (!q) throw new Error("no query");
			return q;
		},
		run: async (fn) => {
			await act(async () => fn(builder.b));
			return builder.b.selectedNodeId as NodeId;
		},
		xml: () => generateFetchXml(builder.query),
		errors: () =>
			validateQueryRules(builder.query)
				.filter((f) => f.severity === "error")
				.map(({ nodeId, attribute, message }) => ({ nodeId, attribute, message })),
	};
	await builder.run((b) => b.setEntity(rootEntity));
	return builder;
}

/** What the tree's "Link-entity" Add item does, then what picking a relationship writes (LinkEntityEditor) */
async function addLink(
	t: Builder,
	parentId: NodeId,
	link: { name: string; from: string; to: string; alias?: string; linkType?: string },
): Promise<NodeId> {
	const id = await t.run((b) => b.addLinkEntity(parentId));
	await t.run((b) =>
		b.updateNode(id, {
			name: link.name,
			from: link.from,
			to: link.to,
			relationshipType: "N1",
			...(link.alias !== undefined ? { alias: link.alias } : {}),
			...(link.linkType ? { linkType: link.linkType } : {}),
		}),
	);
	return id;
}

interface Decoration {
	attribute: string;
	order: { attribute: string; descending?: boolean };
	condition: { attribute: string; operator: string; value?: unknown };
	subCondition: { attribute: string; operator: string; value?: unknown };
}

/** At one level: an attribute, an order, a filter with a condition, and an "or" sub-filter with a condition */
async function decorate(t: Builder, containerId: NodeId, d: Decoration) {
	const attr = await t.run((b) => b.addAttribute(containerId));
	await t.run((b) => b.updateNode(attr, { name: d.attribute }));
	const order = await t.run((b) => b.addOrder(containerId));
	await t.run((b) => b.updateNode(order, d.order));
	const filter = await t.run((b) => b.addFilter(containerId));
	const cond = await t.run((b) => b.addCondition(filter));
	await t.run((b) => b.updateNode(cond, d.condition));
	const sub = await t.run((b) => b.addSubfilter(filter));
	await t.run((b) => b.updateNode(sub, { conjunction: "or" }));
	const subCond = await t.run((b) => b.addCondition(sub));
	await t.run((b) => b.updateNode(subCond, d.subCondition));
	return { attr, order, filter, cond, sub, subCond };
}

const linkById = (q: FetchNode, id: NodeId) => collectAllLinkEntities(q).find((l) => l.id === id);

/** XML → tree → XML */
const reparse = (xml: string): string => {
	const parsed = parseFetchXml(xml);
	if (!parsed.success || !parsed.fetchNode)
		throw new Error(parsed.errors.map((e) => e.message).join("; "));
	return generateFetchXml(parsed.fetchNode);
};

// ---------------------------------------------------------------------------------------------
// 1. Root → four nested link-entities, each with attributes, an order and nested filters
// ---------------------------------------------------------------------------------------------

describe("1. building four nested link-entities from an empty query", () => {
	async function buildDeep() {
		const t = await startBuilder("account");
		const root = t.query.entity.id;
		await decorate(t, root, {
			attribute: "name",
			order: { attribute: "name" },
			condition: { attribute: "statecode", operator: "eq", value: 0 },
			subCondition: { attribute: "name", operator: "like", value: "A%" },
		});
		const l1 = await addLink(t, root, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
			alias: "pc",
		});
		await decorate(t, l1, {
			attribute: "fullname",
			order: { attribute: "lastname", descending: true },
			condition: { attribute: "statecode", operator: "eq", value: 0 },
			subCondition: { attribute: "emailaddress1", operator: "not-null" },
		});
		const l2 = await addLink(t, l1, {
			name: "systemuser",
			from: "systemuserid",
			to: "owninguser",
			alias: "owner",
			linkType: "outer",
		});
		await decorate(t, l2, {
			attribute: "fullname",
			order: { attribute: "fullname" },
			condition: { attribute: "isdisabled", operator: "eq", value: false },
			subCondition: { attribute: "domainname", operator: "ends-with", value: "@contoso.com" },
		});
		const l3 = await addLink(t, l2, {
			name: "businessunit",
			from: "businessunitid",
			to: "businessunitid",
			alias: "bu",
		});
		await decorate(t, l3, {
			attribute: "name",
			order: { attribute: "name" },
			condition: { attribute: "isdisabled", operator: "eq", value: false },
			subCondition: { attribute: "name", operator: "begins-with", value: "Sales" },
		});
		const l4 = await addLink(t, l3, {
			name: "organization",
			from: "organizationid",
			to: "organizationid",
			alias: "org",
			linkType: "matchfirstrowusingcrossapply",
		});
		const attr = await t.run((b) => b.addAttribute(l4));
		await t.run((b) => b.updateNode(attr, { name: "name" }));
		const filter = await t.run((b) => b.addFilter(l4));
		const cond = await t.run((b) => b.addCondition(filter));
		await t.run((b) => b.updateNode(cond, { attribute: "name", operator: "not-null" }));
		const sub = await t.run((b) => b.addSubfilter(filter));
		const subCond = await t.run((b) => b.addCondition(sub));
		await t.run((b) =>
			b.updateNode(subCond, { attribute: "languagecode", operator: "eq", value: 1033 }),
		);
		return { t, l1, l2, l3, l4 };
	}

	it("writes every element inside the right link-entity, with from / to / alias / link-type", async () => {
		const { t } = await buildDeep();
		expect(t.xml()).toBe(
			[
				'<fetch count="50">',
				'  <entity name="account">',
				'    <attribute name="name" />',
				'    <order attribute="name" />',
				'    <filter type="and">',
				'      <condition attribute="statecode" operator="eq" value="0" />',
				'      <filter type="or">',
				'        <condition attribute="name" operator="like" value="A%" />',
				"      </filter>",
				"    </filter>",
				'    <link-entity name="contact" from="contactid" to="primarycontactid" alias="pc">',
				'      <attribute name="fullname" />',
				'      <order attribute="lastname" descending="true" />',
				'      <filter type="and">',
				'        <condition attribute="statecode" operator="eq" value="0" />',
				'        <filter type="or">',
				'          <condition attribute="emailaddress1" operator="not-null" />',
				"        </filter>",
				"      </filter>",
				'      <link-entity name="systemuser" from="systemuserid" to="owninguser" alias="owner" link-type="outer">',
				'        <attribute name="fullname" />',
				'        <order attribute="fullname" />',
				'        <filter type="and">',
				'          <condition attribute="isdisabled" operator="eq" value="false" />',
				'          <filter type="or">',
				'            <condition attribute="domainname" operator="ends-with" value="@contoso.com" />',
				"          </filter>",
				"        </filter>",
				'        <link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="bu">',
				'          <attribute name="name" />',
				'          <order attribute="name" />',
				'          <filter type="and">',
				'            <condition attribute="isdisabled" operator="eq" value="false" />',
				'            <filter type="or">',
				'              <condition attribute="name" operator="begins-with" value="Sales" />',
				"            </filter>",
				"          </filter>",
				'          <link-entity name="organization" from="organizationid" to="organizationid" alias="org" link-type="matchfirstrowusingcrossapply">',
				'            <attribute name="name" />',
				'            <filter type="and">',
				'              <condition attribute="name" operator="not-null" />',
				'              <filter type="and">',
				'                <condition attribute="languagecode" operator="eq" value="1033" />',
				"              </filter>",
				"            </filter>",
				"          </link-entity>",
				"        </link-entity>",
				"      </link-entity>",
				"    </link-entity>",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
	});

	it("has no query-rule errors, so nothing in the tree is flagged", async () => {
		const { t } = await buildDeep();
		expect(t.errors()).toEqual([]);
		expect(getExecutionErrors(parseFetchXml(t.xml()))).toEqual([]);
	});

	it("resolves each link-entity's parent (the table its `to` belongs to) at every depth (#345)", async () => {
		const { t, l1, l2, l3, l4 } = await buildDeep();
		expect(findParentEntityForNode(t.query, l1)).toBe("account");
		expect(findParentEntityForNode(t.query, l2)).toBe("contact");
		expect(findParentEntityForNode(t.query, l3)).toBe("systemuser");
		expect(findParentEntityForNode(t.query, l4)).toBe("businessunit");
	});

	it("round-trips: tree → XML → tree → XML is unchanged, and stable on a second pass", async () => {
		const { t } = await buildDeep();
		const once = reparse(t.xml());
		expect(once).toBe(t.xml());
		expect(reparse(once)).toBe(once);
	});

	it("selects each added node, so the properties panel opens on it", async () => {
		const t = await startBuilder("account");
		const link = await t.run((b) => b.addLinkEntity(t.query.entity.id));
		expect(t.b.selectedNode?.type).toBe("link-entity");
		const nested = await t.run((b) => b.addLinkEntity(link));
		expect(t.b.selectedNodeId).toBe(nested);
		const filter = await t.run((b) => b.addFilter(nested));
		expect(t.b.selectedNode?.type).toBe("filter");
		await t.run((b) => b.addCondition(filter));
		expect(t.b.selectedNode?.type).toBe("condition");
	});
});

// ---------------------------------------------------------------------------------------------
// 2. Link-entities inside filters (any / not any / all / not all), and links nested in them
// ---------------------------------------------------------------------------------------------

describe("2. link-entities inside filters", () => {
	async function buildFilterLinks() {
		const t = await startBuilder("account");
		const root = t.query.entity.id;
		const filter = await t.run((b) => b.addFilter(root));
		await t.run((b) => b.updateNode(filter, { conjunction: "or" }));
		// The filter's "Link-entity (any/not any/all)" Add item
		const anyLink = await t.run((b) => b.addLinkEntity(filter));
		await t.run((b) =>
			b.updateNode(anyLink, {
				name: "contact",
				from: "parentcustomerid",
				to: "accountid",
				alias: "anycontact",
			}),
		);
		const anyFilter = await t.run((b) => b.addFilter(anyLink));
		const anyCond = await t.run((b) => b.addCondition(anyFilter));
		await t.run((b) =>
			b.updateNode(anyCond, { attribute: "statecode", operator: "eq", value: 0 }),
		);
		// A link-entity nested in the "any" link-entity (joined, default inner)
		const nested = await addLink(t, anyLink, {
			name: "systemuser",
			from: "systemuserid",
			to: "owninguser",
			alias: "contactowner",
		});
		const nestedFilter = await t.run((b) => b.addFilter(nested));
		const nestedCond = await t.run((b) => b.addCondition(nestedFilter));
		await t.run((b) =>
			b.updateNode(nestedCond, { attribute: "isdisabled", operator: "eq", value: false }),
		);
		const notAny = await t.run((b) => b.addLinkEntity(filter));
		await t.run((b) =>
			b.updateNode(notAny, {
				name: "opportunity",
				from: "parentaccountid",
				to: "accountid",
				alias: "noopp",
				linkType: "not any",
			}),
		);
		const sub = await t.run((b) => b.addSubfilter(filter));
		const all = await t.run((b) => b.addLinkEntity(sub));
		await t.run((b) =>
			b.updateNode(all, {
				name: "task",
				from: "regardingobjectid",
				to: "accountid",
				alias: "alltasks",
				linkType: "all",
			}),
		);
		const allFilter = await t.run((b) => b.addFilter(all));
		const allCond = await t.run((b) => b.addCondition(allFilter));
		await t.run((b) =>
			b.updateNode(allCond, { attribute: "statecode", operator: "ne", value: 1 }),
		);
		const notAll = await t.run((b) => b.addLinkEntity(sub));
		await t.run((b) =>
			b.updateNode(notAll, {
				name: "email",
				from: "regardingobjectid",
				to: "accountid",
				alias: "notallemails",
				linkType: "not all",
			}),
		);
		return { t, filter, anyLink, nested, notAny, all, notAll };
	}

	it("a link-entity added to a filter defaults to link-type any; one added inside it defaults to inner", async () => {
		const t = await startBuilder("account");
		const filter = await t.run((b) => b.addFilter(t.query.entity.id));
		const anyLink = await t.run((b) => b.addLinkEntity(filter));
		expect(linkById(t.query, anyLink)?.linkType).toBe("any");
		const nested = await t.run((b) => b.addLinkEntity(anyLink));
		expect(linkById(t.query, nested)?.linkType).toBe("inner");
	});

	it("writes the filter link-entities inside their filters, after conditions and sub-filters", async () => {
		const { t } = await buildFilterLinks();
		expect(t.xml()).toBe(
			[
				'<fetch count="50">',
				'  <entity name="account">',
				'    <filter type="or">',
				'      <filter type="and">',
				'        <link-entity name="task" from="regardingobjectid" to="accountid" alias="alltasks" link-type="all">',
				'          <filter type="and">',
				'            <condition attribute="statecode" operator="ne" value="1" />',
				"          </filter>",
				"        </link-entity>",
				'        <link-entity name="email" from="regardingobjectid" to="accountid" alias="notallemails" link-type="not all">',
				"        </link-entity>",
				"      </filter>",
				'      <link-entity name="contact" from="parentcustomerid" to="accountid" alias="anycontact" link-type="any">',
				'        <filter type="and">',
				'          <condition attribute="statecode" operator="eq" value="0" />',
				"        </filter>",
				'        <link-entity name="systemuser" from="systemuserid" to="owninguser" alias="contactowner">',
				'          <filter type="and">',
				'            <condition attribute="isdisabled" operator="eq" value="false" />',
				"          </filter>",
				"        </link-entity>",
				"      </link-entity>",
				'      <link-entity name="opportunity" from="parentaccountid" to="accountid" alias="noopp" link-type="not any">',
				"      </link-entity>",
				"    </filter>",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
	});

	it("has no rule errors, and the parser agrees (nothing blocks Execute)", async () => {
		const { t } = await buildFilterLinks();
		expect(t.errors()).toEqual([]);
		expect(getExecutionErrors(parseFetchXml(t.xml()))).toEqual([]);
	});

	it("flags a joined link type inside a filter on that link-entity (0x8004111c)", async () => {
		const { t, notAny } = await buildFilterLinks();
		await t.run((b) => b.updateNode(notAny, { linkType: "inner" }));
		expect(t.errors()).toEqual([
			expect.objectContaining({ nodeId: notAny, attribute: "link-type" }),
		]);
	});

	it("flags attributes and orders on an any link-entity (0x80041121) on that link-entity", async () => {
		const { t, anyLink } = await buildFilterLinks();
		const attr = await t.run((b) => b.addAttribute(anyLink));
		await t.run((b) => b.updateNode(attr, { name: "fullname" }));
		expect(t.errors()).toEqual([
			expect.objectContaining({ nodeId: anyLink, attribute: "link-type" }),
		]);
	});

	it("resolves the parent of filter link-entities and of the link nested in one (#345)", async () => {
		const { t, anyLink, nested, all } = await buildFilterLinks();
		expect(findParentEntityForNode(t.query, anyLink)).toBe("account");
		expect(findParentEntityForNode(t.query, all)).toBe("account");
		expect(findParentEntityForNode(t.query, nested)).toBe("contact");
	});

	it("round-trips unchanged", async () => {
		const { t } = await buildFilterLinks();
		expect(reparse(t.xml())).toBe(t.xml());
	});

	it("removing a filter removes its link-entities and everything in them", async () => {
		const { t, filter } = await buildFilterLinks();
		await t.run((b) => b.removeNode(filter));
		expect(collectAllLinkEntities(t.query)).toEqual([]);
		expect(t.xml()).toBe(
			['<fetch count="50">', '  <entity name="account">', "  </entity>", "</fetch>"].join(
				"\n",
			),
		);
	});
});

// ---------------------------------------------------------------------------------------------
// 3. Editing at depth 2 to 3
// ---------------------------------------------------------------------------------------------

describe("3. editing nested link-entities", () => {
	/** account → pc (contact) → owner (systemuser) → bu (businessunit), referenced from the root */
	async function buildForEdits() {
		const t = await startBuilder("account");
		const root = t.query.entity.id;
		const l1 = await addLink(t, root, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
			alias: "pc",
		});
		const l2 = await addLink(t, l1, {
			name: "systemuser",
			from: "systemuserid",
			to: "owninguser",
			alias: "owner",
		});
		const ownerAttr = await t.run((b) => b.addAttribute(l2));
		await t.run((b) => b.updateNode(ownerAttr, { name: "fullname" }));
		const l3 = await addLink(t, l2, {
			name: "businessunit",
			from: "businessunitid",
			to: "businessunitid",
			alias: "bu",
		});
		// Root filter conditions on the nested link-entities (entityname), in a sub-filter too
		const filter = await t.run((b) => b.addFilter(root));
		const onOwner = await t.run((b) => b.addCondition(filter));
		await t.run((b) =>
			b.updateNode(onOwner, {
				attribute: "isdisabled",
				operator: "eq",
				value: false,
				entityname: "owner",
			}),
		);
		const sub = await t.run((b) => b.addSubfilter(filter));
		const onBu = await t.run((b) => b.addCondition(sub));
		await t.run((b) =>
			b.updateNode(onBu, {
				attribute: "name",
				operator: "eq",
				value: "Sales",
				entityname: "bu",
			}),
		);
		// A root order on the nested link-entity (entityname)
		const order = await t.run((b) => b.addOrder(root));
		await t.run((b) => b.updateNode(order, { attribute: "fullname", entityname: "owner" }));
		return { t, l1, l2, l3, onOwner, onBu, order };
	}

	it("the starting query is valid", async () => {
		const { t } = await buildForEdits();
		expect(t.errors()).toEqual([]);
	});

	it("an alias change that leaves references behind flags the conditions and orders (0x80041108)", async () => {
		const { t, l2, onOwner, order } = await buildForEdits();
		// A store-level change; the Alias field also updates the references (PropertiesPanel.linkEntity.test.ts)
		await t.run((b) => b.updateNode(l2, { alias: "usr" }));
		const flagged = t.errors().map((e) => e.nodeId);
		expect(flagged).toContain(order);
		expect(flagged).toContain(onOwner);
	});

	it("removing a middle link-entity flags root conditions that pointed into it", async () => {
		const { t, l2, l3, onBu, onOwner } = await buildForEdits();
		await t.run((b) => b.removeNode(l2));
		// The link-entity and its descendants are gone
		expect(linkById(t.query, l2)).toBeUndefined();
		expect(linkById(t.query, l3)).toBeUndefined();
		expect(collectAllLinkEntities(t.query).map((l) => l.alias)).toEqual(["pc"]);
		// Root conditions still say entityname="owner" / "bu", so they're flagged
		const flagged = t.errors().map((e) => e.nodeId);
		expect(flagged).toEqual(expect.arrayContaining([onOwner, onBu]));
	});

	it("removing a middle link-entity flags the root order that pointed into it", async () => {
		const { t, l2, order } = await buildForEdits();
		await t.run((b) => b.removeNode(l2));
		expect(t.errors()).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ nodeId: order, attribute: "entityname" }),
			]),
		);
	});

	it("removing the link-entity that holds the selected node clears the properties panel", async () => {
		const { t, l2, l3 } = await buildForEdits();
		await t.run((b) => b.selectNode(l3));
		await t.run((b) => b.removeNode(l2));
		expect(t.b.selectedNode).toBeNull();
	});

	it("changing a nested link-entity's table by typing keeps from / to / attributes as typed (no metadata check)", async () => {
		const { t, l2 } = await buildForEdits();
		// LinkEntityEditor's Entity field: onUpdate({ name })
		await t.run((b) => b.updateNode(l2, { name: "team" }));
		const link = linkById(t.query, l2)!;
		expect({ name: link.name, from: link.from, to: link.to, alias: link.alias }).toEqual({
			name: "team",
			from: "systemuserid",
			to: "owninguser",
			alias: "owner",
		});
		expect(link.attributes.map((a) => a.name)).toEqual(["fullname"]);
		// The nested link-entity now joins from team: its picker and To list follow
		expect(findParentEntityForNode(t.query, link.links[0].id)).toBe("team");
	});

	it("each link type change is written as link-type, and non-returning types flag the attributes", async () => {
		const { t, l2 } = await buildForEdits();
		const linkTypeOf = () =>
			/<link-entity name="systemuser"[^>]*>/
				.exec(t.xml())?.[0]
				.match(/link-type="([^"]+)"/)?.[1];
		const flaggedOnLink = () =>
			t.errors().some((e) => e.nodeId === l2 && e.attribute === "link-type");
		const expected: [string, string | undefined, boolean][] = [
			["outer", "outer", false],
			["inner", undefined, false],
			["any", "any", true],
			["exists", "exists", true],
			["in", "in", true],
			["matchfirstrowusingcrossapply", "matchfirstrowusingcrossapply", false],
			["not any", "not any", true],
			["all", "all", true],
			["not all", "not all", true],
		];
		for (const [linkType, written, flagged] of expected) {
			await t.run((b) => b.updateNode(l2, { linkType }));
			expect([linkType, linkTypeOf(), flaggedOnLink()]).toEqual([linkType, written, flagged]);
		}
	});

	it("'Change and remove' to a non-returning type keeps the nested link-entity and filters", async () => {
		const { t, l2, l3 } = await buildForEdits();
		// LinkEntityEditor confirmLinkTypeChange(true)
		await t.run((b) =>
			b.updateNode(l2, {
				linkType: "exists",
				attributes: [],
				allAttributes: undefined,
				orders: [],
			}),
		);
		expect(linkById(t.query, l3)).toBeDefined();
		expect(t.errors().filter((e) => e.nodeId === l2)).toEqual([]);
	});
});

// ---------------------------------------------------------------------------------------------
// 4. Round trip for shapes the builder can produce
// ---------------------------------------------------------------------------------------------

describe("4. round trip of builder-produced shapes", () => {
	it("a link-entity with no alias, blank from / to, the same table twice and duplicate aliases", async () => {
		const t = await startBuilder("account");
		const root = t.query.entity.id;
		const a = await addLink(t, root, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
		});
		// Alias cleared in the editor: sanitizeAlias("") || undefined
		await t.run((b) => b.updateNode(a, { alias: undefined }));
		const b2 = await addLink(t, root, {
			name: "contact",
			from: "parentcustomerid",
			to: "accountid",
			alias: "dup",
		});
		const c = await addLink(t, b2, { name: "team", from: "", to: "", alias: "dup" });
		await t.run((b) => b.updateNode(c, { from: undefined, to: undefined }));
		const xml = t.xml();
		expect(xml).toContain(
			'<link-entity name="contact" from="contactid" to="primarycontactid">',
		);
		expect(xml).toContain('<link-entity name="team" alias="dup">');
		expect(reparse(xml)).toBe(xml);
		expect(t.errors()).toEqual([expect.objectContaining({ nodeId: c, attribute: "alias" })]);
	});

	it("placeholder nodes straight from the Add menus (new_entity, new_attribute) round-trip", async () => {
		const t = await startBuilder("account");
		const link = await t.run((b) => b.addLinkEntity(t.query.entity.id));
		await t.run((b) => b.addAttribute(link));
		await t.run((b) => b.addOrder(link));
		const nested = await t.run((b) => b.addLinkEntity(link));
		const filter = await t.run((b) => b.addFilter(nested));
		await t.run((b) => b.addCondition(filter));
		const filterLink = await t.run((b) => b.addLinkEntity(filter));
		expect(linkById(t.query, filterLink)?.linkType).toBe("any");
		expect(reparse(t.xml())).toBe(t.xml());
	});

	it("an aggregate query with groupby / aggregates in a nested link-entity round-trips", async () => {
		const t = await buildAggregate();
		expect(reparse(t.xml())).toBe(t.xml());
	});
});

// ---------------------------------------------------------------------------------------------
// 5. Aggregate queries with link-entities
// ---------------------------------------------------------------------------------------------

async function buildAggregate() {
	const t = await startBuilder("account");
	await t.run((b) => b.updateNode(t.query.id, { aggregate: true, count: undefined }));
	const root = t.query.entity.id;
	const rootAttr = await t.run((b) => b.addAttribute(root));
	await t.run((b) =>
		b.updateNode(rootAttr, { name: "industrycode", groupby: true, alias: "industry" }),
	);
	const l1 = await addLink(t, root, {
		name: "contact",
		from: "contactid",
		to: "primarycontactid",
		alias: "pc",
	});
	const l2 = await addLink(t, l1, {
		name: "systemuser",
		from: "systemuserid",
		to: "owninguser",
		alias: "owner",
	});
	const ownerGroup = await t.run((b) => b.addAttribute(l2));
	await t.run((b) =>
		b.updateNode(ownerGroup, { name: "fullname", groupby: true, alias: "ownername" }),
	);
	const ownerCount = await t.run((b) => b.addAttribute(l2));
	await t.run((b) =>
		b.updateNode(ownerCount, {
			name: "systemuserid",
			aggregate: "countcolumn",
			distinct: true,
			alias: "owners",
		}),
	);
	// An order by alias inside the nested link-entity, and a root order by alias with entityname
	const inner = await t.run((b) => b.addOrder(l2));
	await t.run((b) => b.updateNode(inner, { attribute: undefined, alias: "owners" }));
	const rootOrder = await t.run((b) => b.addOrder(root));
	await t.run((b) =>
		b.updateNode(rootOrder, { attribute: undefined, alias: "ownername", entityname: "owner" }),
	);
	return Object.assign(t, { l2, inner, rootOrder, ownerGroup });
}

describe("5. aggregate queries with nested link-entities", () => {
	it("writes groupby / aggregate attributes in the nested link-entity and orders by alias", async () => {
		const t = await buildAggregate();
		expect(t.xml()).toBe(
			[
				'<fetch aggregate="true">',
				'  <entity name="account">',
				'    <attribute name="industrycode" alias="industry" groupby="true" />',
				'    <order alias="ownername" entityname="owner" />',
				'    <link-entity name="contact" from="contactid" to="primarycontactid" alias="pc">',
				'      <link-entity name="systemuser" from="systemuserid" to="owninguser" alias="owner">',
				'        <attribute name="fullname" alias="ownername" groupby="true" />',
				'        <attribute name="systemuserid" alias="owners" aggregate="countcolumn" distinct="true" />',
				'        <order alias="owners" />',
				"      </link-entity>",
				"    </link-entity>",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
		expect(t.errors()).toEqual([]);
	});

	it("an order by an alias of another link-entity is flagged on the order", async () => {
		const t = await buildAggregate();
		await t.run((b) => b.updateNode(t.inner, { alias: "industry" }));
		expect(t.errors()).toEqual([
			expect.objectContaining({ nodeId: t.inner, attribute: "alias" }),
		]);
	});

	it("renaming an attribute alias finds the orders to update in the nested link-entity and at the root", async () => {
		const t = await buildAggregate();
		const l2 = linkById(t.query, t.l2)! as LinkEntityNode;
		expect(findOrdersUsingAlias(t.query, l2, "ownername").map((o) => o.id)).toEqual([
			t.rootOrder,
		]);
		expect(findOrdersUsingAlias(t.query, l2, "owners").map((o) => o.id)).toEqual([t.inner]);
	});
});

// ---------------------------------------------------------------------------------------------
// 6. Defaults of a new link-entity, and alias uniqueness across the whole tree
// ---------------------------------------------------------------------------------------------

describe("6. defaults and unique aliases", () => {
	it("a link-entity added under a nested link-entity starts as a placeholder with a unique alias", async () => {
		const t = await startBuilder("account");
		const l1 = await t.run((b) => b.addLinkEntity(t.query.entity.id));
		const l2 = await t.run((b) => b.addLinkEntity(l1));
		const l3 = await t.run((b) => b.addLinkEntity(l2));
		const link = linkById(t.query, l3)!;
		expect({
			name: link.name,
			from: link.from,
			to: link.to,
			linkType: link.linkType,
			alias: link.alias,
		}).toEqual({
			name: "new_entity",
			from: "id_field",
			to: "parent_id_field",
			linkType: "inner",
			alias: "new_entity_b1",
		});
	});

	it("auto-generated aliases are unique across joined and filter link-entities", async () => {
		const t = await startBuilder("account");
		const root = t.query.entity.id;
		const l1 = await t.run((b) => b.addLinkEntity(root));
		const filter = await t.run((b) => b.addFilter(l1));
		await t.run((b) => b.addLinkEntity(filter));
		const rootFilter = await t.run((b) => b.addFilter(root));
		const sub = await t.run((b) => b.addSubfilter(rootFilter));
		await t.run((b) => b.addLinkEntity(sub));
		await t.run((b) => b.addLinkEntity(l1));
		const aliases = collectAllLinkEntities(t.query).map((l) => l.alias);
		expect(aliases).toHaveLength(4);
		expect(new Set(aliases).size).toBe(4);
		expect(t.errors().filter((e) => e.attribute === "alias")).toEqual([]);
	});

	it("addLinkEntityWithConfig (many-to-many join) picks an alias unique across the tree", async () => {
		const t = await startBuilder("systemuser");
		const root = t.query.entity.id;
		await addLink(t, root, { name: "team", from: "teamid", to: "x", alias: "team" });
		const intersect = await addLink(t, root, {
			name: "teammembership",
			from: "systemuserid",
			to: "systemuserid",
		});
		let target = "";
		await t.run((b) => {
			target = b.addLinkEntityWithConfig(
				intersect,
				"team",
				"teamid",
				"teamid",
				"inner",
				"NN",
			);
		});
		expect(linkById(t.query, target)?.alias).toBe("team_a1");
	});
});

// ---------------------------------------------------------------------------------------------
// 7. Validation placement and Run gating
// ---------------------------------------------------------------------------------------------

describe("7. validation on the right node", () => {
	it("an entityname order inside a nested link-entity is flagged on that order (0x80041108)", async () => {
		const t = await startBuilder("account");
		const l1 = await addLink(t, t.query.entity.id, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
			alias: "pc",
		});
		const l2 = await addLink(t, l1, {
			name: "systemuser",
			from: "systemuserid",
			to: "owninguser",
			alias: "owner",
		});
		const order = await t.run((b) => b.addOrder(l2));
		await t.run((b) => b.updateNode(order, { attribute: "fullname", entityname: "pc" }));
		expect(t.errors()).toEqual(
			expect.arrayContaining([
				expect.objectContaining({ nodeId: order, attribute: "entityname" }),
			]),
		);
	});

	it("a valueof in a nested link-entity's filter that reaches a later link-entity is flagged on the condition", async () => {
		const t = await startBuilder("account");
		const l1 = await addLink(t, t.query.entity.id, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
			alias: "pc",
		});
		const l2 = await addLink(t, l1, {
			name: "systemuser",
			from: "systemuserid",
			to: "owninguser",
			alias: "owner",
		});
		const filter = await t.run((b) => b.addFilter(l1));
		const cond = await t.run((b) => b.addCondition(filter));
		await t.run((b) =>
			b.updateNode(cond, {
				attribute: "fullname",
				operator: "eq",
				valueof: "owner.fullname",
			}),
		);
		expect(linkById(t.query, l2)).toBeDefined();
		expect(t.errors()).toEqual([
			expect.objectContaining({ nodeId: cond, attribute: "valueof" }),
		]);
	});

	it("a duplicate link alias at depth 3 is flagged on the second link-entity only", async () => {
		const t = await startBuilder("account");
		const l1 = await addLink(t, t.query.entity.id, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
			alias: "x",
		});
		const l2 = await addLink(t, l1, {
			name: "systemuser",
			from: "systemuserid",
			to: "owninguser",
			alias: "y",
		});
		const l3 = await addLink(t, l2, {
			name: "businessunit",
			from: "businessunitid",
			to: "businessunitid",
			alias: "x",
		});
		expect(t.errors()).toEqual([expect.objectContaining({ nodeId: l3, attribute: "alias" })]);
	});
});

// ---------------------------------------------------------------------------------------------
// 8. Code generators (the FetchXML-based ones run locally; C# QueryExpression is converted by
// Dataverse and isn't covered here)
// ---------------------------------------------------------------------------------------------

describe("8. code generators with nested and filter link-entities", () => {
	/** account → pc (contact) → owner (systemuser), plus an any link-entity with a condition to escape */
	async function buildMixed() {
		const t = await startBuilder("account");
		const root = t.query.entity.id;
		const l1 = await addLink(t, root, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
			alias: "pc",
		});
		const l2 = await addLink(t, l1, {
			name: "systemuser",
			from: "systemuserid",
			to: "owninguser",
			alias: "owner",
		});
		const attr = await t.run((b) => b.addAttribute(l2));
		await t.run((b) => b.updateNode(attr, { name: "fullname" }));
		const filter = await t.run((b) => b.addFilter(root));
		const anyLink = await t.run((b) => b.addLinkEntity(filter));
		await t.run((b) =>
			b.updateNode(anyLink, { name: "task", from: "regardingobjectid", to: "accountid" }),
		);
		const anyFilter = await t.run((b) => b.addFilter(anyLink));
		const cond = await t.run((b) => b.addCondition(anyFilter));
		await t.run((b) =>
			b.updateNode(cond, { attribute: "subject", operator: "like", value: "50% & <b> 'x'" }),
		);
		return t;
	}

	it("the Web API URL carries the whole query and decodes back to the same FetchXML", async () => {
		const t = await buildMixed();
		const xml = t.xml();
		const code = generateFxsSyncCode("webapi", xml, "accounts", "https://org.crm.dynamics.com");
		const url =
			/^https:\/\/org\.crm\.dynamics\.com\/api\/data\/v9\.2\/accounts\?fetchXml=(\S+)$/m.exec(
				code,
			);
		expect(url).not.toBeNull();
		const decoded = decodeURIComponent(url![1]);
		expect(decoded).toBe(compactFetchXml(xml));
		expect(reparse(decoded)).toBe(xml);
	});

	it("Power Automate takes the root entity's set name, not a link-entity's", async () => {
		const t = await startBuilder("account");
		const l1 = await addLink(t, t.query.entity.id, {
			name: "contact",
			from: "contactid",
			to: "primarycontactid",
			alias: "pc",
		});
		await addLink(t, l1, { name: "systemuser", from: "systemuserid", to: "owninguser" });
		const spec = generatePowerAutomateSpec(t.xml());
		expect(spec.fields.find((f) => f.label === "Table name")?.value).toBe("accounts");
		expect(spec.fields.find((f) => f.label === "Fetch Xml Query")?.value).toBe(t.xml());
	});

	it("pac CLI gets the nested query on one line that parses back to the same FetchXML", async () => {
		const t = await buildMixed();
		const code = generateFxsSyncCode("pac", t.xml());
		const line = /^pac env fetch --xml '(.*)'$/m.exec(code)?.[1];
		expect(line).toBeDefined();
		// bash: '\'' inside single quotes is one quote
		expect(reparse(line!.replaceAll(`'\\''`, "'"))).toBe(t.xml());
	});
});
