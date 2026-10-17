/**
 * Many-to-many join construction (#85). Relationship metadata as returned by the sandbox; the
 * resulting joins were executed live and matched the intersect entity's rows.
 */
import { describe, expect, it } from "vitest";
import { planIntersectEntityJoins, planManyToManyJoin, relationshipKey, type IntersectRelationship, type ManyToManyRelationship } from "./manyToMany";

const PK: Record<string, string> = { systemuser: "systemuserid", team: "teamid", connectionrole: "connectionroleid", appmodule: "appmoduleid", role: "roleid" };
const pk = (e: string) => PK[e];

const teammembership: ManyToManyRelationship = {
	Entity1LogicalName: "team",
	Entity2LogicalName: "systemuser",
	Entity1IntersectAttribute: "teamid",
	Entity2IntersectAttribute: "systemuserid",
	IntersectEntityName: "teammembership",
};

describe("planManyToManyJoin", () => {
	it("matches the documented systemuser → team example", () => {
		expect(planManyToManyJoin("systemuser", "systemuserid", teammembership, pk)).toEqual({
			intersect: { name: "teammembership", from: "systemuserid", to: "systemuserid" },
			target: { name: "team", from: "teamid", to: "teamid" },
		});
	});

	it("works from the other side (team → systemuser)", () => {
		expect(planManyToManyJoin("team", "teamid", teammembership, pk)).toEqual({
			intersect: { name: "teammembership", from: "teamid", to: "teamid" },
			target: { name: "systemuser", from: "systemuserid", to: "systemuserid" },
		});
	});

	it("self-referential: Entity1 side to Entity2 side, using the two different intersect attributes", () => {
		const rel: ManyToManyRelationship = {
			Entity1LogicalName: "connectionrole",
			Entity2LogicalName: "connectionrole",
			Entity1IntersectAttribute: "connectionroleid",
			Entity2IntersectAttribute: "associatedconnectionroleid",
			IntersectEntityName: "connectionroleassociation",
		};
		expect(planManyToManyJoin("connectionrole", "connectionroleid", rel, pk)).toEqual({
			intersect: { name: "connectionroleassociation", from: "connectionroleid", to: "connectionroleid" },
			target: { name: "connectionrole", from: "connectionroleid", to: "associatedconnectionroleid" },
		});
	});

	it("uses the intersect attributes from metadata, not the related entity's key names", () => {
		const rel: ManyToManyRelationship = {
			Entity1LogicalName: "appmodule",
			Entity2LogicalName: "role",
			Entity1IntersectAttribute: "appmoduleid",
			Entity2IntersectAttribute: "roleid",
			IntersectEntityName: "appmoduleroles",
		};
		expect(planManyToManyJoin("role", "roleid", rel, pk)?.intersect).toEqual({ name: "appmoduleroles", from: "roleid", to: "roleid" });
		expect(planManyToManyJoin("role", "roleid", rel, pk)?.target).toEqual({ name: "appmodule", from: "appmoduleid", to: "appmoduleid" });
	});

	it("returns undefined when the parent isn't part of the relationship", () => {
		expect(planManyToManyJoin("account", "accountid", teammembership, pk)).toBeUndefined();
	});
});

describe("planIntersectEntityJoins (#57)", () => {
	// Sandbox metadata (2026-10-16); every join below ran live
	const systemuserroles: IntersectRelationship = {
		SchemaName: "systemuserroles_association",
		MetadataId: "m1",
		IsValidForAdvancedFind: true,
		Entity1LogicalName: "role",
		Entity2LogicalName: "systemuser",
		Entity1IntersectAttribute: "roleid",
		Entity2IntersectAttribute: "systemuserid",
		IntersectEntityName: "systemuserroles",
	};

	it("joins the intersect entity to each side: from = that side's key, to = the intersect attribute", () => {
		expect(planIntersectEntityJoins("systemuserroles", [systemuserroles], pk)).toEqual([
			{
				SchemaName: "systemuserroles_association",
				RelationshipType: "ManyToOneRelationship",
				MetadataId: "m1:1",
				IsValidForAdvancedFind: true,
				ReferencingEntity: "systemuserroles",
				ReferencingAttribute: "roleid",
				ReferencedEntity: "role",
				ReferencedAttribute: "roleid",
				IntersectSide: 1,
			},
			{
				SchemaName: "systemuserroles_association",
				RelationshipType: "ManyToOneRelationship",
				MetadataId: "m1:2",
				IsValidForAdvancedFind: true,
				ReferencingEntity: "systemuserroles",
				ReferencingAttribute: "systemuserid",
				ReferencedEntity: "systemuser",
				ReferencedAttribute: "systemuserid",
				IntersectSide: 2,
			},
		]);
	});

	it("keeps polymorphic intersect attributes (listmember.entityid → contact.contactid)", () => {
		const listcontact: IntersectRelationship = {
			SchemaName: "listcontact_association",
			MetadataId: "m2",
			Entity1LogicalName: "list",
			Entity2LogicalName: "contact",
			Entity1IntersectAttribute: "listid",
			Entity2IntersectAttribute: "entityid",
			IntersectEntityName: "listmember",
		};
		const joins = planIntersectEntityJoins("listmember", [listcontact], (e) => ({ list: "listid", contact: "contactid" })[e]);
		expect(joins.map((j) => `${j.ReferencedEntity}.${j.ReferencedAttribute} <- ${j.ReferencingAttribute}`)).toEqual([
			"list.listid <- listid",
			"contact.contactid <- entityid",
		]);
	});

	it("self-referential: two joins to the same entity through different attributes", () => {
		const selfRef: IntersectRelationship = {
			SchemaName: "adx_kbarticle_kbarticle",
			MetadataId: "m3",
			Entity1LogicalName: "kbarticle",
			Entity2LogicalName: "kbarticle",
			Entity1IntersectAttribute: "kbarticleidone",
			Entity2IntersectAttribute: "kbarticleidtwo",
			IntersectEntityName: "adx_kbarticle_kbarticle",
		};
		const joins = planIntersectEntityJoins("adx_kbarticle_kbarticle", [selfRef], () => "kbarticleid");
		expect(joins.map((j) => j.ReferencingAttribute)).toEqual(["kbarticleidone", "kbarticleidtwo"]);
		expect(new Set(joins.map(relationshipKey)).size).toBe(2);
	});

	it("ignores relationships of other entities, and sides whose primary key is unknown", () => {
		expect(planIntersectEntityJoins("systemuser", [systemuserroles], pk)).toEqual([]);
		expect(planIntersectEntityJoins("systemuserroles", [systemuserroles], (e) => (e === "role" ? "roleid" : undefined))).toHaveLength(1);
	});
});

describe("relationshipKey", () => {
	it("is the schema name, plus the side for an intersect entity's joins", () => {
		expect(relationshipKey({ SchemaName: "account_primary_contact" })).toBe("account_primary_contact");
		expect(relationshipKey({ SchemaName: "systemuserroles_association", IntersectSide: 2 })).toBe("systemuserroles_association#2");
	});
});
