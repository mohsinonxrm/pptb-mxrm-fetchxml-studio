import { describe, expect, it } from "vitest";
import {
	findLinkAliasReferences,
	findLinkEntityParent,
	findParentEntityForNode,
	getQueryEntities,
	isPlaceholderLinkEntity,
} from "./treeUtils";
import { parseFetchXml } from "./fetchxmlParser";

// Query from issue #54: the link-entity name sorts alphabetically before the root.
const ISSUE_54 = `<fetch aggregate="true">
  <entity name="msdyn_timeentry">
    <attribute name="msdyn_duration" alias="duration" aggregate="sum" />
    <link-entity name="bookableresource" from="bookableresourceid" to="msdyn_bookableresource" alias="BR" />
  </entity>
</fetch>`;

describe("getQueryEntities", () => {
	it("returns the real root even when a link-entity sorts first (#54)", () => {
		const q = getQueryEntities(parseFetchXml(ISSUE_54).fetchNode!);
		expect(q.rootEntityName).toBe("msdyn_timeentry");
		expect(q.entityNames).toEqual(["msdyn_timeentry", "bookableresource"]);
	});

	it("does not reorder its result when the caller sorts a copy", () => {
		const q = getQueryEntities(parseFetchXml(ISSUE_54).fetchNode!);
		[...q.entityNames].sort();
		expect(q.entityNames[0]).toBe("msdyn_timeentry");
	});

	it("changes the key when root and link are swapped", () => {
		const a = getQueryEntities(
			parseFetchXml(`<fetch><entity name="contact"><link-entity name="account" from="accountid" to="parentcustomerid" /></entity></fetch>`).fetchNode!,
		);
		const b = getQueryEntities(
			parseFetchXml(`<fetch><entity name="account"><link-entity name="contact" from="parentcustomerid" to="accountid" /></entity></fetch>`).fetchNode!,
		);
		expect(a.key).not.toBe(b.key);
		expect(a.rootEntityName).toBe("contact");
		expect(b.rootEntityName).toBe("account");
	});

	it("collects nested link-entities once each", () => {
		const q = getQueryEntities(
			parseFetchXml(`<fetch><entity name="systemuser">
				<link-entity name="teammembership" from="systemuserid" to="systemuserid" intersect="true">
					<link-entity name="team" from="teamid" to="teamid" alias="t" />
				</link-entity>
				<link-entity name="team" from="teamid" to="defaultteamid" alias="t2" />
			</entity></fetch>`).fetchNode!,
		);
		expect(q.entityNames).toEqual(["systemuser", "teammembership", "team"]);
	});

	it("handles an empty query", () => {
		expect(getQueryEntities(null)).toEqual({ rootEntityName: null, entityNames: [], key: "" });
	});
});

describe("findParentEntityForNode (#345)", () => {
	const fetch = parseFetchXml(`<fetch>
  <entity name="account">
    <attribute name="name" />
    <link-entity name="contact" from="contactid" to="primarycontactid" alias="pc">
      <attribute name="fullname" />
      <link-entity name="systemuser" from="systemuserid" to="owninguser" alias="owner">
        <attribute name="fullname" />
        <link-entity name="businessunit" from="businessunitid" to="businessunitid" alias="bu" />
      </link-entity>
      <filter>
        <link-entity name="task" from="regardingobjectid" to="contactid" link-type="any" alias="t" />
      </filter>
    </link-entity>
  </entity>
</fetch>`).fetchNode!;
	const contact = fetch.entity.links[0];
	const owner = contact.links[0];
	const parentOf = (id: string) => findParentEntityForNode(fetch, id);

	it("gives a top-level link-entity the root entity", () => {
		expect(parentOf(contact.id)).toBe("account");
	});

	it("gives a nested link-entity the link-entity it sits in, by logical name", () => {
		expect(parentOf(owner.id)).toBe("contact");
		expect(parentOf(owner.links[0].id)).toBe("systemuser");
	});

	it("gives a link-entity in a link-entity's filter that link-entity", () => {
		expect(parentOf(contact.filters[0].links[0].id)).toBe("contact");
	});

	it("gives attributes the entity or link-entity that holds them", () => {
		expect(parentOf(fetch.entity.attributes[0].id)).toBe("account");
		expect(parentOf(contact.attributes[0].id)).toBe("contact");
		expect(parentOf(owner.attributes[0].id)).toBe("systemuser");
	});
});

describe("findLinkEntityParent (#345)", () => {
	const fetch = parseFetchXml(`<fetch>
  <entity name="account">
    <link-entity name="contact" from="contactid" to="primarycontactid" alias="pc">
      <link-entity name="systemuser" from="systemuserid" to="owninguser" alias="owner" />
      <filter>
        <filter>
          <link-entity name="task" from="regardingobjectid" to="contactid" link-type="any" alias="t">
            <link-entity name="systemuser" from="systemuserid" to="ownerid" alias="towner" />
          </link-entity>
        </filter>
      </filter>
    </link-entity>
  </entity>
</fetch>`).fetchNode!;
	const contact = fetch.entity.links[0];
	const task = contact.filters[0].subfilters[0].links[0];

	it("finds the entity or link-entity a link-entity joins from", () => {
		expect(findLinkEntityParent(fetch, contact.id)).toBe(fetch.entity);
		expect(findLinkEntityParent(fetch, contact.links[0].id)).toBe(contact);
	});

	it("finds the owner of a filter's link-entity, and the parent of a link-entity inside it", () => {
		expect(findLinkEntityParent(fetch, task.id)).toBe(contact);
		expect(findLinkEntityParent(fetch, task.links[0].id)).toBe(task);
	});

	it("returns null for a node that isn't a link-entity in the query", () => {
		expect(findLinkEntityParent(fetch, "missing")).toBeNull();
		expect(findLinkEntityParent(null, contact.id)).toBeNull();
	});
});

describe("isPlaceholderLinkEntity (#345)", () => {
	const linkOf = (attrs: string) =>
		parseFetchXml(`<fetch><entity name="account"><link-entity ${attrs} /></entity></fetch>`).fetchNode!.entity
			.links[0];

	it("is true only for the link-entity Add > Link-entity inserts", () => {
		expect(isPlaceholderLinkEntity(linkOf('name="new_entity" from="id_field" to="parent_id_field" alias="new_entity"'))).toBe(true);
	});

	it("is false once any part is set, including a real table called new_entity", () => {
		expect(isPlaceholderLinkEntity(linkOf('name="contact" from="id_field" to="parent_id_field"'))).toBe(false);
		expect(isPlaceholderLinkEntity(linkOf('name="new_entity" from="new_entityid" to="new_entityid"'))).toBe(false);
	});
});

describe("findLinkAliasReferences", () => {
	const fetch = parseFetchXml(`<fetch>
  <entity name="contact">
    <order entityname="acct" attribute="name" />
    <filter>
      <condition entityname="acct" attribute="fax" operator="null" />
      <condition attribute="fullname" operator="eq" valueof="acct.name" />
      <condition entityname="other" attribute="fullname" operator="not-null" />
    </filter>
    <link-entity name="account" from="accountid" to="parentcustomerid" link-type="outer" alias="acct">
      <filter>
        <condition attribute="name" operator="eq" valueof="acct.fax" />
      </filter>
    </link-entity>
    <link-entity name="systemuser" from="systemuserid" to="owninguser" link-type="outer" alias="other" />
  </entity>
</fetch>`).fetchNode!;
	const acct = fetch.entity.links[0];
	const [entityCondition, valueofCondition] = fetch.entity.filters[0].conditions;

	it("finds conditions (entityname, valueof qualifier) and orders that use the alias", () => {
		expect(findLinkAliasReferences(fetch, acct.id, "acct")).toEqual([
			{ nodeId: entityCondition.id, field: "entityname" },
			{ nodeId: valueofCondition.id, field: "valueof", attribute: "name" },
			{ nodeId: acct.filters[0].conditions[0].id, field: "valueof", attribute: "fax" },
			{ nodeId: fetch.entity.orders[0].id, field: "entityname" },
		]);
	});

	it("finds nothing when the alias doesn't name that link-entity alone", () => {
		expect(findLinkAliasReferences(fetch, acct.id, "other")).toEqual([]);
		expect(findLinkAliasReferences(fetch, acct.id, "contact")).toEqual([]);
		expect(findLinkAliasReferences(fetch, acct.id, "")).toEqual([]);
	});
});
