/**
 * Building queries through the UI (#345 review): the tree's context menu and the link-entity properties panel,
 * wired to the real builder the way AppShell wires them. Relationships come from the metadata
 * cache (seeded per test), so the relationship picker offers realistic choices.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { BuilderProvider } from "../state/builderStore";
import { useBuilder, type BuilderContextValue } from "../state/builderContext";
import { TreeView } from "../ui/LeftPane/TreeView";
import { PropertiesPanel } from "../ui/LeftPane/PropertiesPanel";
import { generateFetchXml } from "../model/fetchxml";
import { metadataCache } from "../state/cache";
import { clearMetadataCache } from "../api/dataverseMetadata";
import type { EntityMetadata, RelationshipMetadata } from "../api/pptbClient";
import { cleanup, click, flush, installHost, removeHost, render } from "../../../shared/testing/ui";

afterEach(() => {
	cleanup();
	removeHost();
});

// ---------------------------------------------------------------------------------------------
// Metadata: account ← contact (primarycontactid), contact → systemuser (owninguser),
// systemuser → businessunit, systemuser ↔ team (teammembership)
// ---------------------------------------------------------------------------------------------

const n1 = (
	schema: string,
	referencing: string,
	lookup: string,
	referenced: string,
	key: string,
): RelationshipMetadata => ({
	SchemaName: schema,
	RelationshipType: "OneToManyRelationship",
	ReferencedEntity: referenced,
	ReferencedAttribute: key,
	ReferencingEntity: referencing,
	ReferencingAttribute: lookup,
	MetadataId: schema,
	IsValidForAdvancedFind: true,
});

const accountPrimaryContact = n1(
	"account_primary_contact",
	"account",
	"primarycontactid",
	"contact",
	"contactid",
);
const contactOwningUser = n1(
	"lk_contact_owninguser",
	"contact",
	"owninguser",
	"systemuser",
	"systemuserid",
);
const userBusinessUnit = n1(
	"business_unit_system_users",
	"systemuser",
	"businessunitid",
	"businessunit",
	"businessunitid",
);
const teamMembership: RelationshipMetadata = {
	SchemaName: "teammembership_association",
	RelationshipType: "ManyToManyRelationship",
	ReferencedEntity: "",
	ReferencedAttribute: "",
	ReferencingEntity: "",
	ReferencingAttribute: "",
	MetadataId: "teammembership_association",
	IsValidForAdvancedFind: true,
	Entity1LogicalName: "systemuser",
	Entity2LogicalName: "team",
	Entity1IntersectAttribute: "systemuserid",
	Entity2IntersectAttribute: "teamid",
	IntersectEntityName: "teammembership",
};

const RELATIONSHIPS: Record<
	string,
	Record<"oneToMany" | "manyToOne" | "manyToMany", RelationshipMetadata[]>
> = {
	account: { oneToMany: [], manyToOne: [accountPrimaryContact], manyToMany: [] },
	contact: {
		oneToMany: [accountPrimaryContact],
		manyToOne: [contactOwningUser],
		manyToMany: [],
	},
	systemuser: {
		oneToMany: [contactOwningUser],
		manyToOne: [userBusinessUnit],
		manyToMany: [teamMembership],
	},
	businessunit: { oneToMany: [userBusinessUnit], manyToOne: [], manyToMany: [] },
	team: { oneToMany: [], manyToOne: [], manyToMany: [teamMembership] },
};
const PRIMARY_KEYS: Record<string, string> = {
	account: "accountid",
	contact: "contactid",
	systemuser: "systemuserid",
	businessunit: "businessunitid",
	team: "teamid",
};

/** Seeds the metadata cache, and records the tables whose relationships the picker asks for */
function seedMetadata() {
	clearMetadataCache();
	metadataCache.clear();
	for (const [name, rels] of Object.entries(RELATIONSHIPS))
		metadataCache.setEntityRelationships(`${name}__all`, rels);
	for (const [name, key] of Object.entries(PRIMARY_KEYS))
		metadataCache.setEntityMetadata(name, {
			LogicalName: name,
			PrimaryIdAttribute: key,
		} as EntityMetadata);
	installHost();
	// jsdom has no scrollIntoView; the tree scrolls the selected node into view
	Element.prototype.scrollIntoView = function () {};
}

// ---------------------------------------------------------------------------------------------
// The left pane, wired like AppShell
// ---------------------------------------------------------------------------------------------

let builder: BuilderContextValue | null = null;
function LeftPane() {
	const b = useBuilder();
	builder = b;
	if (!b.fetchQuery) return null;
	return createElement(
		"div",
		null,
		createElement(TreeView, {
			fetchQuery: b.fetchQuery,
			selectedNodeId: b.selectedNodeId,
			onNodeSelect: b.selectNode,
			onAddAttribute: b.addAttribute,
			onAddAllAttributes: b.addAllAttributes,
			onAddOrder: b.addOrder,
			onAddFilter: b.addFilter,
			onAddSubfilter: b.addSubfilter,
			onAddCondition: b.addCondition,
			onAddLinkEntity: b.addLinkEntity,
			onRemoveNode: b.removeNode,
			onUpdateAttributes: b.updateAttributes,
		}),
		createElement(
			"section",
			{ "data-properties": "" },
			createElement(PropertiesPanel, {
				selectedNode: b.selectedNode,
				fetchQuery: b.fetchQuery,
				onNodeUpdate: b.updateNode,
			}),
		),
	);
}

async function startUi(rootEntity: string) {
	seedMetadata();
	builder = null;
	await render(createElement(BuilderProvider, null, createElement(LeftPane)));
	await act(async () => builder!.setEntity(rootEntity));
	await flush();
}

const xml = () => generateFetchXml(builder!.fetchQuery!);

/** The tree row (TreeItemLayout) of a node */
function rowOf(nodeId: string): HTMLElement {
	const item = document.querySelector<HTMLElement>(
		`[data-fui-tree-item-value="${CSS.escape(nodeId)}"]`,
	);
	const layout = item?.querySelector<HTMLElement>(":scope > .fui-TreeItemLayout");
	if (!layout) throw new Error(`no tree row for ${nodeId}`);
	return layout;
}

/** Right-clicks a tree row and returns the context menu's items by label */
async function contextMenu(nodeId: string): Promise<Map<string, HTMLElement>> {
	const row = rowOf(nodeId);
	await act(async () => {
		row.dispatchEvent(
			new MouseEvent("contextmenu", {
				bubbles: true,
				cancelable: true,
				clientX: 5,
				clientY: 5,
			}),
		);
	});
	await flush();
	const items = [...document.querySelectorAll<HTMLElement>('[role="menuitem"]')];
	return new Map(items.map((i) => [(i.textContent ?? "").trim(), i]));
}

async function chooseFromContextMenu(nodeId: string, label: string) {
	const item = (await contextMenu(nodeId)).get(label);
	if (!item) throw new Error(`no "${label}" in the context menu of ${nodeId}`);
	await click(item);
}

/** The relationship picker's options (opens the combobox) */
async function relationshipOptions(): Promise<HTMLElement[]> {
	const panel = document.querySelector("[data-properties]")!;
	for (let i = 0; i < 20 && !panel.querySelector('[role="combobox"]'); i++) await flush();
	const combo = panel.querySelector<HTMLElement>('[role="combobox"]');
	if (!combo) throw new Error(`no relationship picker: ${panel.textContent}`);
	// A click toggles the list: open it only when it's closed
	if (combo.getAttribute("aria-expanded") !== "true") await click(combo);
	return [...document.querySelectorAll<HTMLElement>('[role="option"]')];
}

async function pickRelationship(schemaName: string) {
	const option = (await relationshipOptions()).find((o) => o.textContent?.includes(schemaName));
	if (!option) throw new Error(`no relationship option ${schemaName}`);
	await click(option);
	// The N:N join loads entity metadata before it writes the tree
	await flush(10);
}

const selectedId = () => builder!.selectedNodeId!;

// ---------------------------------------------------------------------------------------------

// UI tests render Fluent menus and pickers: a full parallel run can push one past the default 5 s
describe("UI: building nested link-entities from the tree (#345)", { timeout: 30_000 }, () => {
	it("each level's relationship picker lists its parent link-entity's relationships, and writes from / to", async () => {
		await startUi("account");
		const root = builder!.fetchQuery!.entity.id;

		await chooseFromContextMenu(root, "Link-entity");
		const l1 = selectedId();
		expect((await relationshipOptions()).map((o) => o.textContent)).toEqual([
			expect.stringContaining("account_primary_contact"),
		]);
		await pickRelationship("account_primary_contact");

		await chooseFromContextMenu(l1, "Link-entity");
		const l2 = selectedId();
		// contact's relationships: back to account (1:N) and on to systemuser (N:1)
		const level2 = (await relationshipOptions()).map((o) => o.textContent ?? "");
		expect(level2.some((t) => t.includes("lk_contact_owninguser"))).toBe(true);
		await pickRelationship("lk_contact_owninguser");

		await chooseFromContextMenu(l2, "Link-entity");
		await pickRelationship("business_unit_system_users");

		expect(xml()).toBe(
			[
				'<fetch count="50">',
				'  <entity name="account">',
				'    <link-entity name="contact" from="contactid" to="primarycontactid" alias="contact">',
				'      <link-entity name="systemuser" from="systemuserid" to="owninguser" alias="systemuser">',
				'        <link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="businessunit">',
				"        </link-entity>",
				"      </link-entity>",
				"    </link-entity>",
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
	});

	it("a many-to-many relationship picked at depth 2 adds the intersect entity with the related entity inside it", async () => {
		await startUi("account");
		const root = builder!.fetchQuery!.entity.id;
		await chooseFromContextMenu(root, "Link-entity");
		await pickRelationship("account_primary_contact");
		const l1 = selectedId();
		await chooseFromContextMenu(l1, "Link-entity");
		await pickRelationship("lk_contact_owninguser");
		const l2 = selectedId();
		await chooseFromContextMenu(l2, "Link-entity");
		await pickRelationship("teammembership_association");

		expect(xml()).toContain(
			[
				'      <link-entity name="systemuser" from="systemuserid" to="owninguser" alias="systemuser">',
				'        <link-entity name="teammembership" from="systemuserid" to="systemuserid" intersect="true">',
				'          <link-entity name="team" from="teamid" to="teamid" alias="team">',
				"          </link-entity>",
				"        </link-entity>",
				"      </link-entity>",
			].join("\n"),
		);
		// The intersect entity and the related entity inside it are one undo step (#192)
		expect(builder!.history.undoLabel).toBe("Edit link-entity");
		await act(async () => builder!.history.undo());
		expect(xml()).not.toContain("teammembership");
		expect(xml()).not.toContain('name="team"');
		expect(builder!.history.undoLabel).toBe("Add link-entity");
	});

	it("the same relationship picked twice gets a second, unique alias", async () => {
		await startUi("account");
		const root = builder!.fetchQuery!.entity.id;
		await chooseFromContextMenu(root, "Link-entity");
		await pickRelationship("account_primary_contact");
		await chooseFromContextMenu(root, "Link-entity");
		await pickRelationship("account_primary_contact");
		const aliases = builder!.fetchQuery!.entity.links.map((l) => l.alias);
		expect(aliases).toEqual(["contact", "contact_a1"]);
	});
});

describe("UI: the Add menus of link-entities inside filters", { timeout: 30_000 }, () => {
	it("an any link-entity offers no attributes or orders; a link-entity nested in it does", async () => {
		await startUi("account");
		const root = builder!.fetchQuery!.entity.id;
		await chooseFromContextMenu(root, "Filter");
		const filter = selectedId();
		const filterMenu = await contextMenu(filter);
		expect([...filterMenu.keys()]).toEqual([
			"Condition",
			"Sub-filter",
			"Link-entity (any/not any/all)",
			"Delete",
		]);
		await click(filterMenu.get("Link-entity (any/not any/all)"));
		const anyLink = selectedId();
		const anyMenu = await contextMenu(anyLink);
		const disabled = (label: string) =>
			anyMenu.get(label)?.getAttribute("aria-disabled") === "true";
		expect(
			["All attributes", "Select attributes…", "Blank attribute", "Order by"].map(disabled),
		).toEqual([true, true, true, true]);
		expect([disabled("Filter"), disabled("Link-entity")]).toEqual([false, false]);
		await click(anyMenu.get("Link-entity"));
		// The nested link-entity is joined (inner), so the menu offers attributes: what Dataverse does
		// with them inside an any link-entity is a live check (see the report)
		const nestedMenu = await contextMenu(selectedId());
		expect(nestedMenu.get("Blank attribute")?.getAttribute("aria-disabled")).not.toBe("true");
	});

	it("the link type dropdown of a filter link-entity offers only any / not any / all / not all", async () => {
		await startUi("account");
		const root = builder!.fetchQuery!.entity.id;
		await chooseFromContextMenu(root, "Filter");
		await chooseFromContextMenu(selectedId(), "Link-entity (any/not any/all)");
		const panel = document.querySelector("[data-properties]")!;
		const dropdowns = [...panel.querySelectorAll<HTMLElement>('[role="combobox"]')];
		const linkType = dropdowns.find((d) => d.textContent?.trim() === "Any");
		if (!linkType) throw new Error("no link type dropdown");
		await click(linkType);
		const options = [...document.querySelectorAll<HTMLElement>('[role="option"]')].map(
			(o) => o.getAttribute("text") ?? o.textContent ?? "",
		);
		expect(options.map((o) => o.split(/(?=[A-Z])/)[0])).toEqual([
			"Any",
			"Not any",
			"All",
			"Not all",
		]);
	});
});
