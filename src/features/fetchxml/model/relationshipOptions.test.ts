import { describe, expect, it } from "vitest";
import { buildRelationshipOptions, filterRelationshipOptions } from "./relationshipOptions";
import type { RelationshipMetadata } from "../api/pptbClient";

const rel = (r: Partial<RelationshipMetadata>): RelationshipMetadata =>
	({ MetadataId: r.SchemaName, ReferencedAttribute: "", ReferencingAttribute: "", ReferencedEntity: "", ReferencingEntity: "", ...r }) as RelationshipMetadata;

const ENTITIES: Record<string, string> = { account: "Account", contact: "Contact", opportunity: "Opportunity" };
const ATTRIBUTES: Record<string, string> = { "account.primarycontactid": "Primary Contact", "contact.parentcustomerid": "Company Name" };
const labels = {
	entity: (name: string) => ENTITIES[name],
	attribute: (entity: string, attribute: string) => ATTRIBUTES[`${entity}.${attribute}`],
};

const options = buildRelationshipOptions(
	"account",
	{
		manyToOne: [
			rel({
				SchemaName: "account_primary_contact",
				ReferencedEntity: "contact",
				ReferencedAttribute: "contactid",
				ReferencingEntity: "account",
				ReferencingAttribute: "primarycontactid",
			}),
		],
		oneToMany: [
			rel({
				SchemaName: "contact_customer_accounts",
				ReferencedEntity: "account",
				ReferencedAttribute: "accountid",
				ReferencingEntity: "contact",
				ReferencingAttribute: "parentcustomerid",
			}),
			rel({
				SchemaName: "opportunity_parent_account",
				ReferencedEntity: "account",
				ReferencedAttribute: "accountid",
				ReferencingEntity: "opportunity",
				ReferencingAttribute: "parentaccountid",
			}),
		],
		manyToMany: [
			rel({
				SchemaName: "accountleads_association",
				Entity1LogicalName: "account",
				Entity2LogicalName: "lead",
				Entity1IntersectAttribute: "accountid",
				Entity2IntersectAttribute: "leadid",
				IntersectEntityName: "accountleads",
			}),
		],
	},
	labels,
);

describe("buildRelationshipOptions (#25)", () => {
	it("shows a many-to-one join by the related entity and this entity's lookup", () => {
		const [n1] = options;
		expect(n1.category).toBe("Many-to-one (N:1)");
		expect(n1.primaryText).toBe("Contact (contact)");
		expect(n1.secondaryText).toBe("via Primary Contact (primarycontactid) · account_primary_contact");
		expect(n1.displayText).toBe("Contact via Primary Contact");
	});

	it("shows a one-to-many join by the related entity and its lookup", () => {
		const c = options.find((o) => o.schemaName === "contact_customer_accounts")!;
		expect(c.category).toBe("One-to-many (1:N)");
		expect(c.primaryText).toBe("Contact (contact)");
		expect(c.secondaryText).toBe("via Company Name (parentcustomerid) · contact_customer_accounts");
		expect(c.displayText).toBe("Contact via Company Name");
	});

	it("falls back to logical names when display names aren't known", () => {
		const o = options.find((x) => x.schemaName === "opportunity_parent_account")!;
		expect(o.secondaryText).toBe("via parentaccountid · opportunity_parent_account");
		expect(o.displayText).toBe("Opportunity via parentaccountid");
		const nn = options.find((x) => x.type === "NN")!;
		expect(nn.primaryText).toBe("lead");
		expect(nn.secondaryText).toBe("via accountleads · accountleads_association");
		expect(nn.category).toBe("Many-to-many (N:N)");
		expect(nn.intersectEntity).toBe("accountleads");
	});

	it("keeps the endpoints the link-entity editor needs", () => {
		const c = options.find((o) => o.schemaName === "contact_customer_accounts")!;
		expect(c).toMatchObject({
			type: "1N",
			referencedEntity: "account",
			referencedAttribute: "accountid",
			referencingEntity: "contact",
			referencingAttribute: "parentcustomerid",
		});
	});
});

describe("filterRelationshipOptions (#25)", () => {
	it("matches the entity, the lookup's display or logical name, or the schema name", () => {
		const schemas = (q: string) => filterRelationshipOptions(options, q).map((o) => o.schemaName);
		expect(schemas("company name")).toEqual(["contact_customer_accounts"]);
		expect(schemas("PRIMARYCONTACTID")).toEqual(["account_primary_contact"]);
		expect(schemas("opportunity_parent")).toEqual(["opportunity_parent_account"]);
		expect(schemas("contact")).toEqual(["account_primary_contact", "contact_customer_accounts"]);
		expect(schemas("  ")).toHaveLength(4);
	});
});
