/**
 * "Advanced Find only" scope (#57). Metadata shapes are the sandbox's (2026-10-16): userquery is
 * hidden from Advanced Find but queryable; IsValidForRead = false attributes fail in every query.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { isAttributeInScope, isEntityInScope, isRelationshipInScope, metadataScope } from "./metadataScope";
import { loadAllEntityAttributes, loadEntityAttributes, loadEntityRelationships } from "../api/dataverseMetadata";
import { metadataCache } from "./cache";

const attr = (name: string, af: boolean, read = true) => ({
	LogicalName: name,
	SchemaName: name,
	AttributeType: "String",
	MetadataId: name,
	IsValidForAdvancedFind: { Value: af },
	IsValidForRead: read,
});

describe("scope rules", () => {
	it("entities: Advanced Find only keeps what Advanced Find shows; off keeps all", () => {
		expect(isEntityInScope({ IsValidForAdvancedFind: false }, true)).toBe(false);
		expect(isEntityInScope({ IsValidForAdvancedFind: false }, false)).toBe(true);
		expect(isEntityInScope({ IsValidForAdvancedFind: true }, true)).toBe(true);
		expect(isEntityInScope({}, true)).toBe(true);
	});

	it("attributes: unreadable ones are never offered, whatever the setting", () => {
		expect(isAttributeInScope(attr("isprivate", false, false), false)).toBe(false);
		expect(isAttributeInScope(attr("isprivate", true, false), true)).toBe(false);
		expect(isAttributeInScope(attr("versionnumber", false), false)).toBe(true);
		expect(isAttributeInScope(attr("versionnumber", false), true)).toBe(false);
		expect(isAttributeInScope(attr("name", true), true)).toBe(true);
	});

	it("relationships follow the setting", () => {
		expect(isRelationshipInScope({ IsValidForAdvancedFind: false }, true)).toBe(false);
		expect(isRelationshipInScope({ IsValidForAdvancedFind: false }, false)).toBe(true);
	});

	it("notifies subscribers when the setting changes, and only then", () => {
		const listener = vi.fn();
		const unsubscribe = metadataScope.subscribe(listener);
		metadataScope.setAdvancedFindOnly(false);
		metadataScope.setAdvancedFindOnly(false);
		metadataScope.setAdvancedFindOnly(true);
		unsubscribe();
		metadataScope.setAdvancedFindOnly(false);
		expect(listener).toHaveBeenCalledTimes(2);
		metadataScope.setAdvancedFindOnly(true);
	});
});

describe("metadata loaders (#57)", () => {
	let queryData: ReturnType<typeof vi.fn>;

	beforeEach(() => {
		metadataCache.clear();
		queryData = vi.fn(async (query: string) => {
			if (query.includes("/Attributes?")) {
				return { value: [attr("name", true), attr("versionnumber", false), attr("isprivate", false, false)] };
			}
			if (query.includes("ManyToManyRelationships")) return { value: [] };
			if (query.includes("OneToManyRelationships")) {
				return {
					value: [
						{ SchemaName: "userquery_owner", IsValidForAdvancedFind: false },
						{ SchemaName: "account_contacts", IsValidForAdvancedFind: true },
					],
				};
			}
			return { value: [] };
		});
		(window as unknown as { dataverseAPI: unknown }).dataverseAPI = { queryData };
	});

	afterEach(() => {
		metadataScope.setAdvancedFindOnly(true);
		delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
	});

	it("fetch every attribute once, without a server-side Advanced Find filter", async () => {
		await loadAllEntityAttributes("account");
		await loadEntityAttributes("account", true);
		await loadEntityAttributes("account", false);
		const attributeQueries = queryData.mock.calls.map(([q]) => q as string).filter((q) => q.includes("/Attributes?"));
		expect(attributeQueries).toHaveLength(1);
		expect(attributeQueries[0]).not.toContain("$filter");
		expect(attributeQueries[0]).toContain("IsValidForRead");
	});

	it("pickers get the scoped list; lookups get every attribute", async () => {
		expect((await loadEntityAttributes("account", true)).map((a) => a.LogicalName)).toEqual(["name"]);
		expect((await loadEntityAttributes("account", false)).map((a) => a.LogicalName)).toEqual(["name", "versionnumber"]);
		expect((await loadAllEntityAttributes("account")).map((a) => a.LogicalName)).toEqual(["isprivate", "name", "versionnumber"]);
	});

	it("default to the current setting", async () => {
		metadataScope.setAdvancedFindOnly(false);
		expect((await loadEntityAttributes("account")).map((a) => a.LogicalName)).toContain("versionnumber");
		expect((await loadEntityRelationships("account")).oneToMany.map((r) => r.SchemaName)).toContain("userquery_owner");
		metadataScope.setAdvancedFindOnly(true);
		expect((await loadEntityRelationships("account")).oneToMany.map((r) => r.SchemaName)).toEqual(["account_contacts"]);
		const relationshipQueries = queryData.mock.calls.map(([q]) => q as string).filter((q) => q.includes("Relationships?"));
		expect(relationshipQueries).toHaveLength(3); // fetched once (1:N, N:1, N:N), filtered locally
		expect(relationshipQueries.every((q) => !q.includes("$filter"))).toBe(true);
	});
});

describe("intersect entities get a join to each side of their many-to-many relationship (#57)", () => {
	beforeEach(() => {
		metadataCache.clear();
		const queryData = vi.fn(async (query: string) => {
			if (query.includes("ManyToManyRelationships")) {
				return {
					value: [
						{
							SchemaName: "systemuserroles_association",
							IsValidForAdvancedFind: true,
							Entity1LogicalName: "role",
							Entity2LogicalName: "systemuser",
							Entity1IntersectAttribute: "roleid",
							Entity2IntersectAttribute: "systemuserid",
							IntersectEntityName: "systemuserroles",
						},
					],
				};
			}
			return { value: [] }; // Dataverse gives intersect entities no 1:N / N:1 for these attributes
		});
		const getEntityMetadata = vi.fn(async (name: string) => ({ LogicalName: name, PrimaryIdAttribute: `${name}id` }));
		(window as unknown as { dataverseAPI: unknown }).dataverseAPI = { queryData, getEntityMetadata };
	});

	afterEach(() => {
		delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
	});

	it("offers role and systemuser as many-to-one joins instead of a join through itself", async () => {
		const relationships = await loadEntityRelationships("systemuserroles", false);
		expect(relationships.manyToMany).toEqual([]);
		expect(relationships.manyToOne.map((r) => `${r.ReferencingAttribute} -> ${r.ReferencedEntity}.${r.ReferencedAttribute}`)).toEqual([
			"roleid -> role.roleid",
			"systemuserid -> systemuser.systemuserid",
		]);
	});

	it("follow the relationship's Advanced Find flag", async () => {
		expect((await loadEntityRelationships("systemuserroles", true)).manyToOne).toHaveLength(2);
	});
});
