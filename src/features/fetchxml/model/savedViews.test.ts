import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { generateFetchXml } from "./fetchxml";
import {
	canOverwriteView,
	getQueryTypeLabel,
	getValidationStatus,
	getViewLoadNotice,
	isUnmodifiedView,
	readManagedPropertyValue,
	sortViewsByQueryType,
} from "./savedViews";

/** FetchXML as Dataverse stores it for a system view: legacy attributes and spelled-out defaults. */
const STORED_VIEW_FETCHXML = `<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="account">
    <attribute name="name" />
    <attribute name="primarycontactid" />
    <attribute name="accountid" />
    <order attribute="name" descending="false" />
    <filter type="and">
      <condition attribute="statecode" operator="eq" value="0" />
    </filter>
    <link-entity name="contact" from="contactid" to="primarycontactid" visible="false" link-type="outer" alias="pc">
      <attribute name="emailaddress1" />
    </link-entity>
  </entity>
</fetch>`;

/** What the builder does on Load View: parse, then keep the generated form for comparison. */
function loadView(xml: string) {
	const result = parseFetchXml(xml);
	if (!result.success || !result.fetchNode) throw new Error("parse failed");
	return { fetchNode: result.fetchNode, originalFetchXml: generateFetchXml(result.fetchNode) };
}

describe("isUnmodifiedView (#93)", () => {
	it("matches an unmodified loaded view, so it runs by its ID", () => {
		const view = loadView(STORED_VIEW_FETCHXML);
		// AppShell executes generateFetchXml(builder.fetchQuery)
		expect(isUnmodifiedView(generateFetchXml(view.fetchNode), view.originalFetchXml)).toBe(true);
	});

	it("never matched the stored text (the bug): the generator drops its defaults", () => {
		const view = loadView(STORED_VIEW_FETCHXML);
		expect(isUnmodifiedView(generateFetchXml(view.fetchNode), STORED_VIEW_FETCHXML)).toBe(false);
	});

	it("doesn't match once the query changes", () => {
		const view = loadView(STORED_VIEW_FETCHXML);
		const modified = loadView(STORED_VIEW_FETCHXML.replace('value="0"', 'value="1"'));
		expect(isUnmodifiedView(generateFetchXml(modified.fetchNode), view.originalFetchXml)).toBe(false);
	});

	it("tells whitespace inside values apart", () => {
		expect(isUnmodifiedView('<condition value="a b" />', '<condition value="ab" />')).toBe(false);
		expect(isUnmodifiedView("<fetch />\n", "<fetch />")).toBe(true);
	});
});

describe("query types (#93)", () => {
	it("labels every view that isn't a public view", () => {
		expect(getQueryTypeLabel(0)).toBeUndefined();
		expect(getQueryTypeLabel(undefined)).toBeUndefined();
		expect(getQueryTypeLabel(1)).toBe("Advanced find");
		expect(getQueryTypeLabel(2)).toBe("Associated");
		expect(getQueryTypeLabel(4)).toBe("Quick find");
		expect(getQueryTypeLabel(64)).toBe("Lookup");
		expect(getQueryTypeLabel(8192)).toBe("Offline template");
		// Unknown types are still listed, with their number
		expect(getQueryTypeLabel(999)).toBe("Query type 999");
	});

	it("sorts public views first (default on top), then the other types", () => {
		const views = [
			{ name: "Quick Find Active Accounts", queryType: 4 },
			{ name: "Unknown", queryType: 999 },
			{ name: "Lookup View", queryType: 64 },
			{ name: "Active Accounts", queryType: 0, isDefault: false },
			{ name: "My Active Accounts", queryType: 0, isDefault: true },
			{ name: "Account Advanced Find View", queryType: 1 },
			{ name: "Accounts: Influenced Deals", queryType: 0 },
			{ name: "Account Associated View", queryType: 2 },
		];
		expect(sortViewsByQueryType(views).map((v) => v.name)).toEqual([
			"My Active Accounts",
			"Accounts: Influenced Deals",
			"Active Accounts",
			"Account Advanced Find View",
			"Account Associated View",
			"Quick Find Active Accounts",
			"Lookup View",
			"Unknown",
		]);
	});
});

describe("saving (#93)", () => {
	it("shows high and critical performance messages as errors", () => {
		expect(getValidationStatus([])).toBe("valid");
		expect(getValidationStatus([{ Severity: 0 }, { Severity: 1 }])).toBe("warning");
		expect(getValidationStatus([{ Severity: 1 }, { Severity: 2 }])).toBe("error");
		expect(getValidationStatus([{ Severity: 3 }])).toBe("error");
	});

	it("blocks overwriting only a system view that isn't customizable", () => {
		expect(canOverwriteView({ type: "system", isCustomizable: false })).toBe(false);
		expect(canOverwriteView({ type: "system", isCustomizable: true })).toBe(true);
		// Unknown: Dataverse decides
		expect(canOverwriteView({ type: "system" })).toBe(true);
		expect(canOverwriteView({ type: "personal", isCustomizable: false })).toBe(true);
	});

	it("reads the iscustomizable managed property", () => {
		expect(
			readManagedPropertyValue({
				Value: false,
				CanBeChanged: false,
				ManagedPropertyLogicalName: "iscustomizableanddeletable",
			}),
		).toBe(false);
		expect(readManagedPropertyValue({ Value: true })).toBe(true);
		expect(readManagedPropertyValue(true)).toBe(true);
		expect(readManagedPropertyValue(undefined)).toBeUndefined();
		expect(readManagedPropertyValue({})).toBeUndefined();
	});
});

describe("getViewLoadNotice (#93)", () => {
	it("says nothing when the view loads as-is", () => {
		expect(getViewLoadNotice("Active Accounts", parseFetchXml(STORED_VIEW_FETCHXML))).toBeNull();
	});

	it("reports a view whose FetchXML can't be parsed", () => {
		const notice = getViewLoadNotice("Broken", parseFetchXml("<fetch><entity name='account'>"));
		expect(notice?.type).toBe("error");
		expect(notice?.body).toContain('"Broken"');
	});

	it("lists what the tree changed", () => {
		const result = parseFetchXml(
			`<fetch><entity name="account"><attribute name="name" alias="1bad" /></entity></fetch>`,
		);
		const notice = getViewLoadNotice("Aliased", result);
		expect(notice?.type).toBe("warning");
		expect(notice?.body).toContain("<attribute>");
		expect(notice?.body).toContain("Saving the view saves these changes.");
	});
});
