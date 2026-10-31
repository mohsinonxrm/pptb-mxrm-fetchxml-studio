import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { deleteAllUnavailableReason, deleteSelectionUnavailableReason, MAX_DIRECT_DELETE } from "./deleteScope";
import type { FetchNode } from "./nodes";

const q = (xml: string): FetchNode => {
	const r = parseFetchXml(xml);
	if (!r.success || !r.fetchNode) throw new Error(`parse failed: ${xml}`);
	return r.fetchNode;
};

const RAN = `<fetch><entity name="account"><attribute name="name"/><order attribute="name"/>
  <filter><condition attribute="statecode" operator="eq" value="1"/></filter></entity></fetch>`;

/** The query that ran vs the tree now: is "Delete all" allowed? (#174) */
const allowed = (executed: string | null, current: string) =>
	deleteAllUnavailableReason(executed ? q(executed) : null, q(current)) === undefined;

describe("Delete all uses the query on screen (#174)", () => {
	it("is allowed when the tree is the query that ran", () => {
		expect(allowed(RAN, RAN)).toBe(true);
	});

	it("is blocked before any run", () => {
		expect(deleteAllUnavailableReason(null, q(RAN))).toMatch(/^Run the query first/);
	});

	describe("left-pane edits that change which records match: blocked until re-run", () => {
		it.each([
			["a condition's value", RAN.replace('value="1"', 'value="0"')],
			["a condition's operator", RAN.replace('operator="eq"', 'operator="ne"')],
			["the filter removed", RAN.replace(/<filter>.*<\/filter>/s, "")],
			["a condition added", RAN.replace("</filter>", '<condition attribute="name" operator="like" value="a%"/></filter>')],
			["an inner link-entity", RAN.replace("</entity>", '<link-entity name="contact" from="parentcustomerid" to="accountid" link-type="inner"/></entity>')],
			["an outer link-entity with a filter", RAN.replace("</entity>", '<link-entity name="contact" from="parentcustomerid" to="accountid" link-type="outer"><filter><condition attribute="firstname" operator="eq" value="x"/></filter></link-entity></entity>')],
			["top added", RAN.replace("<fetch>", '<fetch top="10">')],
		])("%s", (_label, current) => {
			expect(deleteAllUnavailableReason(q(RAN), q(current))).toMatch(/^The query has changed/);
		});
	});

	describe("results-grid edits that keep the same records: still allowed", () => {
		it.each([
			["a column added", RAN.replace('<attribute name="name"/>', '<attribute name="name"/><attribute name="telephone1"/>')],
			["a column removed", RAN.replace('<attribute name="name"/>', "")],
			["all columns", RAN.replace('<attribute name="name"/>', "<all-attributes/>")],
			["sorted by another column", RAN.replace('<order attribute="name"/>', '<order attribute="createdon" descending="true"/>')],
			["the sort removed", RAN.replace('<order attribute="name"/>', "")],
			["related columns (outer link-entity, columns only)", RAN.replace("</entity>", '<link-entity name="contact" from="contactid" to="primarycontactid" link-type="outer" alias="pc"><attribute name="fullname"/></link-entity></entity>')],
		])("%s", (_label, current) => {
			expect(allowed(RAN, current)).toBe(true);
		});
	});

	it("with top, the sort decides which records are in the first N: a new sort needs a re-run", () => {
		const topRan = RAN.replace("<fetch>", '<fetch top="10">');
		expect(allowed(topRan, topRan.replace('<order attribute="name"/>', '<order attribute="createdon"/>'))).toBe(false);
	});
});

describe("Delete of more than 100 selected records needs prvBulkDelete (#177)", () => {
	it("allows up to 100 selected records without the privilege", () => {
		expect(deleteSelectionUnavailableReason(1, false)).toBeUndefined();
		expect(deleteSelectionUnavailableReason(MAX_DIRECT_DELETE, false)).toBeUndefined();
	});

	it("blocks more than 100 without the privilege, and names it", () => {
		const reason = deleteSelectionUnavailableReason(MAX_DIRECT_DELETE + 1, false);
		expect(reason).toContain("prvBulkDelete");
		expect(reason).toContain("100 or fewer");
	});

	it("allows any number with the privilege (a Bulk Delete job)", () => {
		expect(deleteSelectionUnavailableReason(5000, true)).toBeUndefined();
	});
});
