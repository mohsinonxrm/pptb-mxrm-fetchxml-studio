/**
 * Link-entity editing from the Properties panel, found in the #345 review: a link-entity under one
 * that isn't set up yet, renaming an alias that conditions and orders use, and "compare to" the
 * root entity from a self-join.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { PropertiesPanel } from "./PropertiesPanel";
import { ValueOfPicker } from "./PropertiesPanel/editors/ValueOfPicker";
import { BuilderProvider } from "../../state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../state/builderContext";
import { clearMetadataCache } from "../../api/dataverseMetadata";
import { metadataCache } from "../../state/cache";
import type { FetchNode } from "../../model/nodes";
import { cleanup, installHost, removeHost, render, typeInto } from "../../../../shared/testing/ui";

afterEach(() => {
	cleanup();
	removeHost();
});

/** Stub host that records which tables' relationships were asked for */
function recordRelationshipRequests(): string[] {
	clearMetadataCache();
	metadataCache.clear();
	const tables: string[] = [];
	installHost({
		dataverseAPI: {
			queryData: async (query: unknown) => {
				const match = /^EntityDefinitions\(LogicalName='([^']+)'\)\/\w+Relationships/.exec(String(query));
				if (match) tables.push(match[1]);
				return { value: [] };
			},
		},
	});
	return tables;
}

/** Wires PropertiesPanel to real builder state, like AppShell does */
let builder: BuilderContextValue;
function Harness() {
	const b = useBuilder();
	builder = b;
	return createElement(PropertiesPanel, { selectedNode: b.selectedNode, fetchQuery: b.fetchQuery, onNodeUpdate: b.updateNode });
}

async function loadAndSelect(xml: string, pick: (fetch: FetchNode) => { id: string }) {
	await render(createElement(BuilderProvider, null, createElement(Harness)));
	await act(async () => builder.loadFetchXml(xml));
	const id = pick(builder.fetchQuery!).id;
	await act(async () => builder.selectNode(id));
}

/** The link-entity editor's Alias input */
const aliasInput = () => document.querySelector<HTMLInputElement>('input[placeholder="e.g. parent_account"]') ?? undefined;

describe("link-entity under a link-entity that isn't set up yet (#345)", () => {
	const xml = `<fetch>
		<entity name="account">
			<link-entity name="new_entity" from="id_field" to="parent_id_field" alias="new_entity">
				<link-entity name="new_entity" from="id_field" to="parent_id_field" alias="new_entity1" />
			</link-entity>
		</entity>
	</fetch>`;

	it("doesn't ask Dataverse for the placeholder table's relationships, and says what to do", async () => {
		const tables = recordRelationshipRequests();
		await loadAndSelect(xml, (fetch) => fetch.entity.links[0].links[0]);
		expect(tables).not.toContain("new_entity");
		expect(document.body.textContent).toContain("Set up the link-entity this one is inside first");
	});

	it("still offers the root's relationships for the top-level placeholder", async () => {
		const tables = recordRelationshipRequests();
		await loadAndSelect(xml, (fetch) => fetch.entity.links[0]);
		expect(new Set(tables)).toEqual(new Set(["account"]));
	});
});

describe("renaming a link-entity's alias", () => {
	const xml = `<fetch>
		<entity name="contact">
			<attribute name="fullname" />
			<order entityname="acct" attribute="name" />
			<filter>
				<condition entityname="acct" attribute="fax" operator="null" />
				<condition attribute="fullname" operator="eq" valueof="acct.name" />
				<condition entityname="a" attribute="fullname" operator="not-null" />
			</filter>
			<link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer" alias="acct">
				<attribute name="name" />
			</link-entity>
			<link-entity name="systemuser" from="systemuserid" to="owninguser" link-type="outer" alias="a" />
		</entity>
	</fetch>`;
	const select = (fetch: FetchNode) => fetch.entity.links[0];

	it("updates the conditions (entityname, valueof) and orders that use it", async () => {
		recordRelationshipRequests();
		await loadAndSelect(xml, select);
		await typeInto(aliasInput(), "parent");

		const fetch = builder.fetchQuery!;
		expect(fetch.entity.links[0].alias).toBe("parent");
		expect(fetch.entity.filters[0].conditions[0].entityname).toBe("parent");
		expect(fetch.entity.filters[0].conditions[1].valueof).toBe("parent.name");
		expect(fetch.entity.orders[0].entityname).toBe("parent");
	});

	it("keeps following the same references when the alias briefly matches another one", async () => {
		recordRelationshipRequests();
		await loadAndSelect(xml, select);
		const alias = aliasInput();
		await typeInto(alias, "a"); // the systemuser link-entity's alias, while typing
		await typeInto(alias, "ab");

		const conditions = builder.fetchQuery!.entity.filters[0].conditions;
		expect(conditions[0].entityname).toBe("ab");
		expect(conditions[1].valueof).toBe("ab.name");
		expect(conditions[2].entityname).toBe("a"); // still the systemuser link-entity
	});

	it("leaves the references alone while the alias is empty, then follows the new one", async () => {
		recordRelationshipRequests();
		await loadAndSelect(xml, select);
		const alias = aliasInput();
		await typeInto(alias, "");
		expect(builder.fetchQuery!.entity.links[0].alias).toBeUndefined();
		expect(builder.fetchQuery!.entity.filters[0].conditions[0].entityname).toBe("acct");
		await typeInto(alias, "p");
		expect(builder.fetchQuery!.entity.filters[0].conditions[0].entityname).toBe("p");
	});
});

describe("compare to the root entity (ValueOfPicker)", () => {
	const renderPicker = (onRootRecord: boolean) =>
		render(
			createElement(ValueOfPicker, {
				fetchQuery: null,
				rootEntityName: "account",
				effectiveEntityName: "account",
				onRootRecord,
				scope: { rootName: "account", aliases: ["parent"], selfAlias: "parent", laterAliases: [] },
				sameAlias: "parent",
				valueof: "account.name",
				compatibleTypes: ["String"],
				onChange: () => {},
			}),
		);

	it("offers the root from a self-join (a link-entity to the root's own entity)", async () => {
		installHost();
		await renderPicker(false);
		expect(document.querySelector('button[role="combobox"]')?.textContent).toContain("(root entity)");
	});

	it("doesn't list the root twice for a condition on the root record", async () => {
		installHost();
		await renderPicker(true);
		expect(document.querySelector('button[role="combobox"]')?.textContent).not.toContain("(root entity)");
	});
});

describe("a condition's entityname naming a link-entity by its logical name (#345, verified live)", () => {
	it("loads that link-entity's attributes, not the root's", async () => {
		clearMetadataCache();
		metadataCache.clear();
		const tables: string[] = [];
		installHost({
			dataverseAPI: {
				queryData: async (query: unknown) => {
					const match = /^EntityDefinitions\(LogicalName='([^']+)'\)\/Attributes/.exec(String(query));
					if (match) tables.push(match[1]);
					return { value: [] };
				},
			},
		});
		await loadAndSelect(
			`<fetch>
				<entity name="contact">
					<attribute name="fullname" />
					<filter><condition entityname="account" attribute="fax" operator="null" /></filter>
					<link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer" alias="a" />
				</entity>
			</fetch>`,
			(fetch) => fetch.entity.filters[0].conditions[0],
		);
		expect(tables).toContain("account");
		expect(tables).not.toContain("contact");
		expect(document.querySelector('button[role="combobox"]')?.textContent).toContain("a (account)");
	});
});
