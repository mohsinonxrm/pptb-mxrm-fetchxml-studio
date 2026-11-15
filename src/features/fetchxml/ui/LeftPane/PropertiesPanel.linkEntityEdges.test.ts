/**
 * Link-entity editing edge cases from the 1.5.1 review (#345): a placeholder parent whose To
 * attribute was picked first, and entityname naming the root entity or a link-entity by its
 * logical name.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { PropertiesPanel } from "./PropertiesPanel";
import { BuilderProvider } from "../../state/builderStore";
import { useBuilder, type BuilderContextValue } from "../../state/builderContext";
import { clearMetadataCache } from "../../api/dataverseMetadata";
import { metadataCache } from "../../state/cache";
import type { FetchNode } from "../../model/nodes";
import { validateQueryRules } from "../../model/fetchxmlRules";
import { parseFetchXml } from "../../model/fetchxmlParser";
import { cleanup, installHost, removeHost, render } from "../../../../shared/testing/ui";

afterEach(() => {
	cleanup();
	removeHost();
});

function recordRequests(pattern: RegExp): string[] {
	clearMetadataCache();
	metadataCache.clear();
	const tables: string[] = [];
	installHost({
		dataverseAPI: {
			queryData: async (query: unknown) => {
				const match = pattern.exec(String(query));
				if (match) tables.push(match[1]);
				return { value: [] };
			},
		},
	});
	return tables;
}

let builder: BuilderContextValue;
function Harness() {
	const b = useBuilder();
	builder = b;
	return createElement(PropertiesPanel, {
		selectedNode: b.selectedNode,
		fetchQuery: b.fetchQuery,
		onNodeUpdate: b.updateNode,
	});
}

async function loadAndSelect(xml: string, pick: (fetch: FetchNode) => { id: string }) {
	await render(createElement(BuilderProvider, null, createElement(Harness)));
	await act(async () => builder.loadFetchXml(xml));
	const id = pick(builder.fetchQuery!).id;
	await act(async () => builder.selectNode(id));
}

describe("placeholder parent whose To attribute was picked first", () => {
	// Add > Link-entity, then the user picks the placeholder's To attribute (the root's attribute
	// picker works) before choosing its entity. The parent is still `new_entity` (its own editor
	// still shows the relationship picker, LinkEntityEditor shouldShowPicker), but
	// isPlaceholderLinkEntity requires from AND to to be the placeholder values too.
	const xml = `<fetch>
		<entity name="account">
			<link-entity name="new_entity" from="id_field" to="primarycontactid" alias="new_entity">
				<link-entity name="new_entity" from="id_field" to="parent_id_field" alias="new_entity1" />
			</link-entity>
		</entity>
	</fetch>`;

	it("the child doesn't ask Dataverse for the relationships of the placeholder table new_entity", async () => {
		const tables = recordRequests(
			/^EntityDefinitions\(LogicalName='([^']+)'\)\/\w+Relationships/,
		);
		await loadAndSelect(xml, (fetch) => fetch.entity.links[0].links[0]);
		expect(tables).not.toContain("new_entity");
	});
});

describe("entityname = the root entity's name in a self-join", () => {
	// Root account, one joined account link-entity (alias parent). The rule (fetchxmlRules
	// unknownEntityname) resolves entityname="account" to the root entity first; the condition
	// editor's new by-name fallback resolves it to the link-entity instead.
	const xml = `<fetch>
		<entity name="account">
			<attribute name="name" />
			<filter><condition entityname="account" attribute="name" operator="not-null" /></filter>
			<link-entity name="account" from="accountid" to="parentaccountid" link-type="outer" alias="parent" />
		</entity>
	</fetch>`;

	it("the rule accepts it as the root entity", () => {
		const findings = validateQueryRules(parseFetchXml(xml).fetchNode!).filter(
			(f) => f.attribute === "entityname",
		);
		expect(findings).toEqual([]);
	});

	it("the condition editor shows it as the root entity too, as the rule reads it", async () => {
		recordRequests(/^EntityDefinitions\(LogicalName='([^']+)'\)\/Attributes/);
		await loadAndSelect(xml, (fetch) => fetch.entity.filters[0].conditions[0]);
		const entityDropdown = document.querySelector('button[role="combobox"]');
		expect(entityDropdown?.textContent).not.toContain("parent (account)");
	});
});

describe("entityname naming an aliased link-entity by its logical name", () => {
	// The by-name fallback makes the Entity dropdown show "a (account)", but selectedOptions stays
	// [node.entityname] = ["account"], and no option has that value (options use ref.identifier, the
	// alias). Nothing is selected; picking "a (account)" to confirm it clears attribute/operator/value.
	it("the Entity dropdown marks the link-entity it shows as selected", async () => {
		recordRequests(/^EntityDefinitions\(LogicalName='([^']+)'\)\/Attributes/);
		await loadAndSelect(
			`<fetch>
				<entity name="contact">
					<filter><condition entityname="account" attribute="fax" operator="null" /></filter>
					<link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer" alias="a" />
				</entity>
			</fetch>`,
			(fetch) => fetch.entity.filters[0].conditions[0],
		);
		const dropdown = document.querySelector<HTMLElement>('button[role="combobox"]');
		expect(dropdown?.textContent).toContain("a (account)");
		await act(async () => dropdown!.click());
		const selected = [
			...document.querySelectorAll('[role="option"][aria-selected="true"]'),
		].map((o) => o.textContent);
		expect(selected).toEqual(["a (account)"]);
	});
});
