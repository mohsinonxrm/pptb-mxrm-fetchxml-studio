/**
 * useRootEntityMetadata (#301): the query's entity metadata loads when the root entity (or the set of
 * entities) changes and not otherwise, a response for an earlier root is dropped (#54), a failed
 * load is logged and leaves that part empty, and after a connection change (AppShell clears the
 * caches and mounts the tool again, #66) everything loads afresh from the new environment.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import {
	cleanup,
	deferred,
	flush,
	installHost,
	removeHost,
	render,
	type HostStub,
} from "../../shared/testing/ui";
import { createFakeDataverse, sampleTables } from "../../shared/testing/fakeDataverse";
import {
	loadEntityMetadata,
	loadEntityRelationships,
} from "../../features/fetchxml/api/dataverseMetadata";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { clearMetadataLoadErrors } from "../../features/fetchxml/ui/LeftPane/PropertiesPanel/metadataLoadErrors";
import type { EntityMetadata, RelationshipMetadata } from "../../features/fetchxml/api/pptbClient";
import type { FetchNode } from "../../features/fetchxml/model/nodes";
import {
	useRootEntityMetadata,
	type RootEntityMetadata,
	type RootEntityMetadataInputs,
} from "./useRootEntityMetadata";

let latest: RootEntityMetadata;

function Probe(props: RootEntityMetadataInputs) {
	latest = useRootEntityMetadata(props);
	return null;
}

/** A query rooted at `entity`, with link-entities to `links`; `id` stands for a new tree */
const query = (entity: string, links: string[] = [], id = "e") =>
	({
		id: "f",
		type: "fetch",
		options: {},
		entity: {
			id,
			type: "entity",
			name: entity,
			attributes: [],
			orders: [],
			filters: [],
			links: links.map((name, i) => ({
				id: `l${i}`,
				type: "link-entity",
				name,
				attributes: [],
				orders: [],
				filters: [],
				links: [],
			})),
		},
	}) as unknown as FetchNode;

const relationship = (schemaName: string) => ({ SchemaName: schemaName }) as RelationshipMetadata;

let loaders: {
	loadEntityMetadata: Mock<RootEntityMetadataInputs["loadEntityMetadata"]>;
	loadRelationships: Mock<RootEntityMetadataInputs["loadRelationships"]>;
};

const show = (fetchQuery: FetchNode | null, advancedFindOnly = false) =>
	render(createElement(Probe, { fetchQuery, advancedFindOnly, ...loaders }));

const attributeNames = (entity: string) => [
	...(latest.attributeMetadata.get(entity)?.keys() ?? []),
];

let dataverse: ReturnType<typeof createFakeDataverse>;
let host: HostStub;

/** An in-memory Dataverse with these tables, behind a stubbed host */
function connect(tables = sampleTables()) {
	dataverse = createFakeDataverse(tables);
	host = installHost({ dataverseAPI: dataverse.api as unknown as HostStub["dataverseAPI"] });
}

beforeEach(() => {
	metadataCache.clear();
	clearMetadataLoadErrors();
	connect();
	// The real loaders (through metadataCache), counted
	loaders = {
		loadEntityMetadata:
			vi.fn<RootEntityMetadataInputs["loadEntityMetadata"]>(loadEntityMetadata),
		loadRelationships:
			vi.fn<RootEntityMetadataInputs["loadRelationships"]>(loadEntityRelationships),
	};
	vi.spyOn(console, "error").mockImplementation(() => undefined);
});

afterEach(() => {
	cleanup();
	removeHost();
	metadataCache.clear();
	clearMetadataLoadErrors();
	vi.restoreAllMocks();
});

describe("useRootEntityMetadata: loading", () => {
	it("with no root entity, nothing loads and everything is empty", async () => {
		await show(null);
		expect(latest.attributeMetadata.size).toBe(0);
		expect(latest.entityMetadata).toBeNull();
		expect(latest.lookupRelationships).toEqual([]);
		expect(latest.oneToManyRelationships).toEqual([]);
		expect(latest.isLoadingRelationships).toBe(false);
		expect(loaders.loadEntityMetadata).not.toHaveBeenCalled();
		expect(loaders.loadRelationships).not.toHaveBeenCalled();
	});

	it("loads the attributes of every entity in the query, and the root's metadata and relationships", async () => {
		await show(query("account", ["contact"]));
		expect([...latest.attributeMetadata.keys()]).toEqual(["account", "contact"]);
		expect(attributeNames("account")).toEqual(["accountid", "name", "statecode", "statuscode"]);
		expect(attributeNames("contact")).toEqual(["contactid", "fullname", "statecode"]);
		expect(latest.entityMetadata?.LogicalName).toBe("account");
		expect(latest.entityMetadata?.PrimaryIdAttribute).toBe("accountid");
		expect(loaders.loadEntityMetadata.mock.calls).toEqual([["account"]]);
		expect(loaders.loadRelationships.mock.calls).toEqual([["account"]]);
		expect(latest.isLoadingRelationships).toBe(false);
	});

	it("keeps the relationships the loader returns, and shows the loading flag while they load", async () => {
		const pending = deferred<{
			oneToMany: RelationshipMetadata[];
			manyToOne: RelationshipMetadata[];
		}>();
		loaders.loadRelationships.mockReturnValue(pending.promise);
		await show(query("account"));
		expect(latest.isLoadingRelationships).toBe(true);
		pending.resolve({
			manyToOne: [relationship("account_primary_contact")],
			oneToMany: [relationship("contact_customer_accounts")],
		});
		await flush();
		expect(latest.isLoadingRelationships).toBe(false);
		expect(latest.lookupRelationships).toEqual([relationship("account_primary_contact")]);
		expect(latest.oneToManyRelationships).toEqual([relationship("contact_customer_accounts")]);
	});

	it("loads again when the root entity changes, and never shows the old root's metadata for the new one", async () => {
		await show(query("account"));
		expect(latest.entityMetadata?.LogicalName).toBe("account");

		const contact = deferred<EntityMetadata>();
		loaders.loadEntityMetadata.mockReturnValueOnce(contact.promise);
		await show(query("contact", [], "e2"));
		expect(latest.entityMetadata).toBeNull();
		expect([...latest.attributeMetadata.keys()]).toEqual(["contact"]);
		expect(loaders.loadEntityMetadata.mock.calls).toEqual([["account"], ["contact"]]);
		expect(loaders.loadRelationships.mock.calls).toEqual([["account"], ["contact"]]);

		contact.resolve(await loadEntityMetadata("contact"));
		await flush();
		expect(latest.entityMetadata?.LogicalName).toBe("contact");
	});

	it("drops a response for a root entity the query no longer has (#54)", async () => {
		const account = deferred<EntityMetadata>();
		const accountRelationships = deferred<{
			oneToMany: RelationshipMetadata[];
			manyToOne: RelationshipMetadata[];
		}>();
		loaders.loadEntityMetadata.mockReturnValueOnce(account.promise);
		loaders.loadRelationships.mockReturnValueOnce(accountRelationships.promise);
		await show(query("account"));
		await show(query("contact", [], "e2"));
		expect(latest.entityMetadata?.LogicalName).toBe("contact");

		account.resolve(await loadEntityMetadata("account"));
		accountRelationships.resolve({
			manyToOne: [relationship("account_primary_contact")],
			oneToMany: [],
		});
		await flush();
		expect(latest.entityMetadata?.LogicalName).toBe("contact");
		expect(latest.lookupRelationships).toEqual([]);
		expect(latest.isLoadingRelationships).toBe(false);
	});

	it("loads again when a link-entity is added, and when Advanced Find only changes (#57)", async () => {
		await show(query("account"));
		await show(query("account", ["contact"]));
		expect([...latest.attributeMetadata.keys()]).toEqual(["account", "contact"]);
		expect(loaders.loadEntityMetadata).toHaveBeenCalledTimes(2);

		await show(query("account", ["contact"]), true);
		expect(loaders.loadRelationships).toHaveBeenCalledTimes(3);
	});

	it("doesn't load again while the root and the set of entities stay the same", async () => {
		await show(query("account", ["contact"]));
		const requests = dataverse.calls.length;
		// A new tree object for the same entities: an attribute or filter edit, or the same entity again
		await show(query("account", ["contact"], "e2"));
		await show(query("account", ["contact", "contact"], "e3"));
		expect(loaders.loadEntityMetadata).toHaveBeenCalledTimes(1);
		expect(loaders.loadRelationships).toHaveBeenCalledTimes(1);
		expect(dataverse.calls.length).toBe(requests);
		expect(latest.entityMetadata?.LogicalName).toBe("account");
	});

	it("clears everything but the relationships when the root entity is removed", async () => {
		loaders.loadRelationships.mockResolvedValue({
			manyToOne: [relationship("account_primary_contact")],
			oneToMany: [],
		});
		await show(query("account"));
		await show(null);
		expect(latest.attributeMetadata.size).toBe(0);
		expect(latest.entityMetadata).toBeNull();
		// As before the move: the relationships are kept until the next root entity loads its own
		expect(latest.lookupRelationships).toEqual([relationship("account_primary_contact")]);
	});
});

describe("useRootEntityMetadata: failures", () => {
	it("logs each failed load and leaves that part empty", async () => {
		const failure = new Error("0x80040216: An unexpected error occurred.");
		host.dataverseAPI.queryData = async (q: unknown) => {
			if (String(q).includes("/Attributes")) throw failure;
			return { value: [] };
		};
		loaders.loadEntityMetadata.mockRejectedValue(failure);
		loaders.loadRelationships.mockRejectedValue(failure);
		await show(query("account"));

		expect([...latest.attributeMetadata.keys()]).toEqual(["account"]);
		expect(latest.attributeMetadata.get("account")?.size).toBe(0);
		expect(latest.entityMetadata).toBeNull();
		expect(latest.lookupRelationships).toEqual([]);
		expect(latest.oneToManyRelationships).toEqual([]);
		expect(latest.isLoadingRelationships).toBe(false);
		expect(console.error).toHaveBeenCalledWith(
			"Failed to load attributes for account:",
			failure,
		);
		expect(console.error).toHaveBeenCalledWith("Failed to load entity metadata:", failure);
		expect(console.error).toHaveBeenCalledWith("Failed to load relationships:", failure);
	});

	it("a link-entity's failed attributes don't stop the others", async () => {
		const failure = new Error("0x80060888: Could not find entity 'contact'");
		const answer = host.dataverseAPI.queryData;
		host.dataverseAPI.queryData = async (q: unknown) => {
			if (String(q).startsWith("EntityDefinitions(LogicalName='contact')/Attributes"))
				throw failure;
			return answer(q);
		};
		await show(query("account", ["contact"]));
		expect(attributeNames("account")).toEqual(["accountid", "name", "statecode", "statuscode"]);
		expect(latest.attributeMetadata.get("contact")?.size).toBe(0);
		expect(latest.entityMetadata?.LogicalName).toBe("account");
		expect(console.error).toHaveBeenCalledWith(
			"Failed to load attributes for contact:",
			failure,
		);
	});

	it("drops the earlier relationships when the next root's relationships fail", async () => {
		loaders.loadRelationships.mockResolvedValueOnce({
			manyToOne: [relationship("account_primary_contact")],
			oneToMany: [relationship("contact_customer_accounts")],
		});
		await show(query("account"));
		loaders.loadRelationships.mockRejectedValueOnce(new Error("failed"));
		await show(query("contact", [], "e2"));
		expect(latest.lookupRelationships).toEqual([]);
		expect(latest.oneToManyRelationships).toEqual([]);
		expect(latest.isLoadingRelationships).toBe(false);
	});
});

describe("useRootEntityMetadata: connection change", () => {
	/** What AppShell does when the connection changes (#66, #201): clear the caches, mount again */
	function changeConnection(tables: ReturnType<typeof sampleTables>) {
		metadataCache.clear();
		clearMetadataLoadErrors();
		cleanup();
		connect(tables);
	}

	it("starts empty and loads the new environment's metadata, not the cached one", async () => {
		await show(query("account"));
		expect(latest.entityMetadata?.ObjectTypeCode).toBe(1);

		// The same entity in the new environment, with other metadata
		const tables = sampleTables();
		tables[0].objectTypeCode = 10001;
		tables[0].attributes = [{ name: "accountid", type: "Uniqueidentifier" }];
		changeConnection(tables);

		const metadata = deferred<EntityMetadata>();
		loaders.loadEntityMetadata.mockReturnValueOnce(metadata.promise);
		await show(query("account"));
		expect(latest.entityMetadata).toBeNull();
		expect(attributeNames("account")).toEqual(["accountid"]);

		metadata.resolve(await loadEntityMetadata("account"));
		await flush();
		expect(latest.entityMetadata?.ObjectTypeCode).toBe(10001);
		expect(dataverse.calls).toContain("getEntityMetadata account");
	});
});
