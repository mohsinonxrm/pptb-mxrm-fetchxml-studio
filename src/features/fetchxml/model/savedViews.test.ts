import { describe, expect, it } from "vitest";
import { parseFetchXml } from "./fetchxmlParser";
import { generateFetchXml } from "./fetchxml";
import {
	applyQuickFindValue,
	canOverwriteView,
	canSetDefaultView,
	countQuickFindPlaceholders,
	defaultViewsNotClearedMessage,
	getQueryTypeLabel,
	getValidationStatus,
	getViewLoadNotice,
	getViewSecondaryText,
	hasWriteAccess,
	isSharedView,
	isUnmodifiedView,
	readManagedPropertyValue,
	sortViewsByQueryType,
	toQuickFindSearchValue,
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

const ME = "AAAAAAAA-0000-0000-0000-000000000001";
const NANCY = "bbbbbbbb-0000-0000-0000-000000000002";

describe("isSharedView and getViewSecondaryText (#142)", () => {
	it("says who shared a personal view the user doesn't own", () => {
		const view = { type: "personal" as const, ownerId: NANCY, ownerName: "Nancy Anderson" };
		expect(isSharedView(view, ME)).toBe(true);
		expect(getViewSecondaryText(view, ME)).toBe("Shared by Nancy Anderson");
	});

	it("has no second line for the user's own public views (IDs compared case-insensitively)", () => {
		const view = { type: "personal" as const, ownerId: ME.toLowerCase(), ownerName: "Me" };
		expect(isSharedView(view, ME)).toBe(false);
		expect(getViewSecondaryText(view, ME)).toBeUndefined();
	});

	it("doesn't call a view shared when the owner or the user is unknown", () => {
		expect(isSharedView({ type: "personal", ownerId: NANCY }, undefined)).toBe(false);
		expect(isSharedView({ type: "personal" }, ME)).toBe(false);
		expect(isSharedView({ type: "system", ownerId: NANCY }, ME)).toBe(false);
	});

	it("falls back to 'Shared with you' without the owner's name", () => {
		expect(getViewSecondaryText({ type: "personal", ownerId: NANCY }, ME)).toBe("Shared with you");
	});

	it("names the query type, or the default public view, for system views", () => {
		expect(getViewSecondaryText({ type: "system", queryType: 4, isDefault: true }, ME)).toBe("Quick find view");
		expect(getViewSecondaryText({ type: "system", queryType: 0, isDefault: true }, ME)).toBe("Default view");
		expect(getViewSecondaryText({ type: "system", queryType: 0 }, ME)).toBeUndefined();
		expect(getViewSecondaryText({ type: "system", queryType: 99 }, ME)).toBe("Query type 99");
	});

	it("combines the owner and the query type", () => {
		const view = { type: "personal" as const, ownerId: NANCY, ownerName: "Nancy", queryType: 1 };
		expect(getViewSecondaryText(view, ME)).toBe("Shared by Nancy · Advanced find view");
	});
});

describe("hasWriteAccess (#142)", () => {
	it("reads the AccessRights string RetrievePrincipalAccess returns (verified live)", () => {
		expect(hasWriteAccess("ReadAccess, WriteAccess, CreateAccess, DeleteAccess, ShareAccess, AssignAccess")).toBe(true);
		expect(hasWriteAccess("ReadAccess")).toBe(false);
		expect(hasWriteAccess("None")).toBe(false);
	});

	it("reads a flags number (WriteAccess = 2)", () => {
		expect(hasWriteAccess(3)).toBe(true);
		expect(hasWriteAccess(1)).toBe(false);
	});

	it("is undefined when it can't be read, so the save is allowed", () => {
		expect(hasWriteAccess(undefined)).toBeUndefined();
		expect(hasWriteAccess({})).toBeUndefined();
	});
});

describe("canSetDefaultView (#142)", () => {
	it("applies to public system views only", () => {
		expect(canSetDefaultView({ type: "system", queryType: 0 })).toBe(true);
		expect(canSetDefaultView({ type: "system" })).toBe(true);
		expect(canSetDefaultView({ type: "system", queryType: 4 })).toBe(false);
		expect(canSetDefaultView({ type: "personal", queryType: 0 })).toBe(false);
	});
});

/** The account quick find view as Dataverse stores it (shape verified live) */
const QUICK_FIND_FETCHXML = `<fetch version="1.0" mapping="logical">
  <entity name="account">
    <attribute name="name" />
    <filter type="and">
      <condition attribute="statecode" operator="eq" value="0" />
    </filter>
    <filter type="or" isquickfindfields="1">
      <condition attribute="name" operator="like" value="{0}" />
      <condition attribute="accountnumber" operator="like" value="{0}" />
      <condition attribute="emailaddress1" operator="like" value='{0}' />
    </filter>
  </entity>
</fetch>`;

describe("quick find test value (#142)", () => {
	it("counts the condition values holding {0}", () => {
		expect(countQuickFindPlaceholders(QUICK_FIND_FETCHXML)).toBe(3);
		expect(countQuickFindPlaceholders(STORED_VIEW_FETCHXML)).toBe(0);
		expect(
			countQuickFindPlaceholders(`<condition attribute="x" operator="in"><value>{0}</value><value>a</value></condition>`),
		).toBe(1);
	});

	it("ignores {0} outside condition values", () => {
		expect(
			countQuickFindPlaceholders(`<fetch><entity name="account"><attribute name="name" alias="{0}" /></entity></fetch>`),
		).toBe(0);
	});

	it("adds the trailing % like requires, unless it's there or turned off", () => {
		expect(toQuickFindSearchValue("contoso", true)).toBe("contoso%");
		expect(toQuickFindSearchValue("con%", true)).toBe("con%");
		expect(toQuickFindSearchValue("contoso", false)).toBe("contoso");
		expect(toQuickFindSearchValue("  ", true)).toBe("");
		expect(toQuickFindSearchValue(" a ", true)).toBe("a%");
	});

	it("replaces {0} in every condition value and leaves the rest of the query alone", () => {
		const run = applyQuickFindValue(QUICK_FIND_FETCHXML, "a%");
		expect(countQuickFindPlaceholders(run)).toBe(0);
		expect(run).toContain(`<condition attribute="name" operator="like" value="a%" />`);
		expect(run).toContain(`<condition attribute="emailaddress1" operator="like" value='a%' />`);
		expect(run).toContain(`<condition attribute="statecode" operator="eq" value="0" />`);
		expect(run.replace(/a%/g, "{0}")).toBe(QUICK_FIND_FETCHXML);
	});

	it("XML-escapes the value for its quotes and elements", () => {
		const run = applyQuickFindValue(QUICK_FIND_FETCHXML, `A&B "x" 'y' <z>%`);
		expect(run).toContain(`value="A&amp;B &quot;x&quot; 'y' &lt;z&gt;%"`);
		expect(run).toContain(`value='A&amp;B "x" &apos;y&apos; &lt;z&gt;%'`);
		const values = applyQuickFindValue(`<condition attribute="x" operator="in"><value>{0}</value></condition>`, `a<b"`);
		expect(values).toBe(`<condition attribute="x" operator="in"><value>a&lt;b"</value></condition>`);
		expect(new DOMParser().parseFromString(run, "application/xml").querySelector("parsererror")).toBeNull();
	});

	it("leaves the generated tree FetchXML of a loaded quick find view runnable", () => {
		const view = loadView(QUICK_FIND_FETCHXML);
		expect(countQuickFindPlaceholders(view.originalFetchXml)).toBe(3);
		expect(countQuickFindPlaceholders(applyQuickFindValue(view.originalFetchXml, "contoso%"))).toBe(0);
	});
});

describe("defaultViewsNotClearedMessage (#307)", () => {
	it("names the view that stayed a default and adds Dataverse's error", () => {
		const message = defaultViewsNotClearedMessage("FXS view", [{ name: "All Territories", error: "No privilege." }]);
		expect(message).toBe(
			`"FXS view" is the default view now, but "All Territories" is still marked as the default too. Set the default again in Power Apps (the entity's Views, then Set as default) to fix it. Details: No privilege.`,
		);
	});

	it("names every view when more than one stayed a default", () => {
		const message = defaultViewsNotClearedMessage("FXS view", [
			{ name: "A", error: "E1." },
			{ name: "B", error: "E2." },
		]);
		expect(message).toContain('"A", "B" are still marked as the default too');
		expect(message).toContain("Details: E1. E2.");
	});

	it("says it couldn't check when the other defaults couldn't be read", () => {
		const message = defaultViewsNotClearedMessage("FXS view", [{ name: "", error: "Network error." }]);
		expect(message).toContain("couldn't check whether another view is still marked as the default too");
		expect(message).toContain("Details: Network error.");
	});
});
