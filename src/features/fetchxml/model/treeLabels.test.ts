import { describe, expect, it } from "vitest";
import { countDescendants, deleteConfirmationText, linkEntityLabel, linkEntityLookup, showsNoAttributesHint } from "./treeLabels";
import { parseFetchXml } from "./fetchxmlParser";
import type { LinkEntityNode } from "./nodes";

const link = (overrides: Partial<LinkEntityNode>): LinkEntityNode => ({
	id: "l1",
	type: "link-entity",
	name: "contact",
	from: "contactid",
	to: "primarycontactid",
	linkType: "outer",
	attributes: [],
	orders: [],
	filters: [],
	links: [],
	...overrides,
});

describe("linkEntityLabel (#25)", () => {
	it("names the parent's lookup (to) for a many-to-one join", () => {
		expect(linkEntityLabel(link({ relationshipType: "N1" }), "accountid")).toBe("contact via primarycontactid (outer)");
	});

	it("names the related entity's lookup (from) for a one-to-many join", () => {
		const l = link({ from: "parentcustomerid", to: "accountid", linkType: "inner", relationshipType: "1N" });
		expect(linkEntityLabel(l)).toBe("contact via parentcustomerid (inner)");
	});

	it("treats a join to the parent's primary key as one-to-many without a relationship type", () => {
		const l = link({ from: "parentcustomerid", to: "accountid" });
		expect(linkEntityLookup(l, "accountid")).toBe("parentcustomerid");
		expect(linkEntityLookup(l, undefined)).toBe("accountid");
		expect(linkEntityLookup(link({}), "accountid")).toBe("primarycontactid");
	});

	it("keeps the alias label unchanged", () => {
		expect(linkEntityLabel(link({ alias: "owner", linkType: "inner" }), "accountid")).toBe("owner (inner)");
	});

	it("falls back to the name when no attribute is set", () => {
		expect(linkEntityLabel(link({ from: "", to: "" }))).toBe("contact (outer)");
		expect(linkEntityLabel(link({ from: "contactid", to: "", relationshipType: "N1" }))).toBe("contact via contactid (outer)");
	});
});

describe("showsNoAttributesHint (#24)", () => {
	it("shows for a returning link type with no attributes", () => {
		expect(showsNoAttributesHint(link({ linkType: "inner" }))).toBe(true);
		expect(showsNoAttributesHint(link({ linkType: "matchfirstrowusingcrossapply" }))).toBe(true);
	});

	it("doesn't show for link types that can't return attributes", () => {
		for (const linkType of ["exists", "in", "any", "not any", "all", "not all"] as const) {
			expect(showsNoAttributesHint(link({ linkType }))).toBe(false);
		}
	});

	it("doesn't show with attributes, all-attributes or for an intersect link-entity", () => {
		expect(showsNoAttributesHint(link({ attributes: [{ id: "a", type: "attribute", name: "fullname" }] }))).toBe(false);
		expect(showsNoAttributesHint(link({ allAttributes: { id: "aa", type: "all-attributes", enabled: true } }))).toBe(false);
		expect(showsNoAttributesHint(link({ allAttributes: { id: "aa", type: "all-attributes", enabled: false } }))).toBe(true);
		expect(showsNoAttributesHint(link({ intersect: true }))).toBe(false);
	});
});

describe("countDescendants (#24)", () => {
	const fetch = parseFetchXml(`<fetch>
	<entity name="account">
		<attribute name="name" />
		<filter>
			<condition attribute="statecode" operator="eq" value="0" />
			<filter type="or">
				<condition attribute="name" operator="like" value="a%" />
			</filter>
			<link-entity name="task" from="regardingobjectid" to="accountid" link-type="not any" />
		</filter>
		<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer">
			<all-attributes />
			<attribute name="fullname" />
			<order attribute="fullname" />
			<link-entity name="systemuser" from="systemuserid" to="owninguser" />
		</link-entity>
		<link-entity name="systemuser" from="systemuserid" to="owninguser" />
	</entity>
</fetch>`).fetchNode!;
	const [filter] = fetch.entity.filters;
	const [contact, empty] = fetch.entity.links;

	it("counts every item at any depth", () => {
		// condition, sub-filter + its condition, link-entity
		expect(countDescendants(filter)).toBe(4);
		// all-attributes, attribute, order, nested link-entity
		expect(countDescendants(contact)).toBe(4);
		expect(countDescendants(fetch.entity)).toBe(1 + 1 + 4 + 1 + 4 + 1);
	});

	it("is 0 for an empty link-entity (deleted without asking)", () => {
		expect(countDescendants(empty)).toBe(0);
	});

	it("words the confirmation for one and many items", () => {
		expect(deleteConfirmationText(1)).toBe("This also deletes the item inside it.");
		expect(deleteConfirmationText(4)).toBe("This also deletes the 4 items inside it.");
	});
});
