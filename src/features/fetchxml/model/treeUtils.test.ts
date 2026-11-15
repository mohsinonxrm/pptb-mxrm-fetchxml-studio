import { describe, expect, it } from "vitest";
import { findParentEntityForNode, getQueryEntities } from "./treeUtils";
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
			parseFetchXml(
				`<fetch><entity name="contact"><link-entity name="account" from="accountid" to="parentcustomerid" /></entity></fetch>`,
			).fetchNode!,
		);
		const b = getQueryEntities(
			parseFetchXml(
				`<fetch><entity name="account"><link-entity name="contact" from="parentcustomerid" to="accountid" /></entity></fetch>`,
			).fetchNode!,
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
