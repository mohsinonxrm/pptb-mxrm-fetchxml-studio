/**
 * #345: a link-entity nested in another link-entity joins from that link-entity, so its
 * relationship picker must list the parent link-entity's relationships, not the root entity's.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { PropertiesPanel } from "./PropertiesPanel";
import { BuilderProvider } from "../../state/builderStore";
import { parseFetchXml } from "../../model/fetchxmlParser";
import { clearMetadataCache } from "../../api/dataverseMetadata";
import { metadataCache } from "../../state/cache";
import type { LinkEntityNode } from "../../model/nodes";
import { cleanup, installHost, removeHost, render } from "../../../../shared/testing/ui";

afterEach(() => {
	cleanup();
	removeHost();
});

/** Renders the panel for a link-entity and returns the tables whose relationships were asked for */
async function relationshipTablesFor(
	xml: string,
	pick: (links: LinkEntityNode[]) => LinkEntityNode,
) {
	clearMetadataCache();
	metadataCache.clear();
	const asked: string[] = [];
	installHost({
		dataverseAPI: {
			queryData: async (query: unknown) => {
				const match = /^EntityDefinitions\(LogicalName='([^']+)'\)\/\w+Relationships/.exec(
					String(query),
				);
				if (match) asked.push(match[1]);
				return { value: [] };
			},
		},
	});
	const fetchQuery = parseFetchXml(xml).fetchNode!;
	const selectedNode = pick(fetchQuery.entity.links);
	await render(
		createElement(
			BuilderProvider,
			null,
			createElement(PropertiesPanel, { selectedNode, fetchQuery, onNodeUpdate: () => {} }),
		),
	);
	return [...new Set(asked)];
}

describe("PropertiesPanel: nested link-entity relationship picker (#345)", () => {
	it("lists the parent link-entity's relationships, not the root entity's", async () => {
		const tables = await relationshipTablesFor(
			`<fetch>
				<entity name="account">
					<link-entity name="contact" from="contactid" to="primarycontactid" alias="pc">
						<link-entity name="new_entity" from="id_field" to="parent_id_field" alias="new_entity" />
					</link-entity>
				</entity>
			</fetch>`,
			(links) => links[0].links[0],
		);
		expect(tables).toEqual(["contact"]);
	});

	it("uses the logical name of a parent two levels down, not its alias", async () => {
		const tables = await relationshipTablesFor(
			`<fetch>
				<entity name="account">
					<link-entity name="contact" from="contactid" to="primarycontactid" alias="pc">
						<link-entity name="systemuser" from="systemuserid" to="owninguser" alias="owner">
							<link-entity name="new_entity" from="id_field" to="parent_id_field" alias="new_entity" />
						</link-entity>
					</link-entity>
				</entity>
			</fetch>`,
			(links) => links[0].links[0].links[0],
		);
		expect(tables).toEqual(["systemuser"]);
	});

	it("still lists the root entity's relationships for a top-level link-entity", async () => {
		const tables = await relationshipTablesFor(
			`<fetch>
				<entity name="account">
					<link-entity name="new_entity" from="id_field" to="parent_id_field" alias="new_entity" />
				</entity>
			</fetch>`,
			(links) => links[0],
		);
		expect(tables).toEqual(["account"]);
	});
});
