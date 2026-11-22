/**
 * Open view on the app bar (#278, owner decision Q2): a popover with a search box over the entity's
 * views, "My views" above "System views", the default view marked and the open view checked. It
 * waits for an entity, loads the views each time it opens, asks before it discards unsaved changes,
 * and works from More once it has overflowed. The view lookups are mocked.
 */
import { act, createElement, useState } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	cleanup,
	click,
	flush,
	render,
	stubOverflowWidths,
	toasts,
	typeInto,
} from "../../../../shared/testing/ui";
import { AppBar, type AppBarProps } from "./AppBar";
import { OverflowToolbar } from "../../../../shared/components/commandBar";
import { openViewDisabledReason, openViewGroups } from "./appBarRules";
import type { SavedView } from "../../api/pptbClient";

const api = vi.hoisted(() => ({ getAllViews: vi.fn(), whoAmI: vi.fn() }));
vi.mock("../../api/pptbClient", async (importOriginal) => ({
	...(await importOriginal<typeof import("../../api/pptbClient")>()),
	...api,
}));

const ME = "aaaaaaaa-0000-0000-0000-000000000001";
const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";
const view = (extra: Partial<SavedView> & Pick<SavedView, "id" | "name" | "type">): SavedView => ({
	fetchxml: XML,
	layoutxml: "<grid/>",
	...extra,
});
const active = view({
	id: "s1",
	name: "Active Accounts",
	type: "system",
	isDefault: true,
	queryType: 0,
});
const inactive = view({ id: "s2", name: "Inactive Accounts", type: "system", queryType: 0 });
const mine = view({ id: "p1", name: "My big accounts", type: "personal", ownerId: ME });
const shared = view({
	id: "p2",
	name: "Team accounts",
	type: "personal",
	ownerId: "bbbbbbbb-0000-0000-0000-000000000002",
	ownerName: "Pat Doe",
});

const barProps = (props: Partial<AppBarProps> = {}): AppBarProps => ({
	run: { onClick: () => {}, busy: false },
	saveView: {
		fetchXml: XML,
		layoutXml: "",
		entityLogicalName: "account",
		objectTypeCode: 1,
		primaryIdAttribute: "accountid",
		targetView: null,
	},
	getCurrentXml: () => XML,
	viewName: null,
	unsaved: false,
	onNewQuery: () => {},
	openView: {
		entityLogicalName: "account",
		entityDisplayName: "Account",
		entitySetName: "accounts",
		onOpen: vi.fn(),
	},
	...props,
});
const renderBar = async (props: Partial<AppBarProps> = {}) => {
	const all = barProps(props);
	await render(createElement(AppBar, all));
	await flush(5);
	return all;
};

const bar = () => document.querySelector<HTMLElement>('[role="toolbar"]');
const openViewButton = () =>
	[...(bar()?.querySelectorAll<HTMLButtonElement>("button") ?? [])].find(
		(b) => b.textContent?.trim() === "Open view",
	);
const surface = () => document.querySelector<HTMLElement>(".fui-PopoverSurface");
const search = () => surface()?.querySelector<HTMLInputElement>('input[aria-label="Search views"]');
const groups = () =>
	[...(surface()?.querySelectorAll<HTMLElement>('[role="group"]') ?? [])].map((g) => ({
		label: g.querySelector('[role="presentation"]')?.textContent?.trim(),
		views: [...g.querySelectorAll<HTMLElement>('[role="option"]')].map(optionText),
	}));
/** An option's name and its second line, space-separated (the check icon has no text) */
function optionText(o: HTMLElement): string {
	const lines = [...o.querySelectorAll("span")].find(
		(s) => s.children.length > 0 && [...s.children].every((c) => c.tagName === "SPAN"),
	);
	return lines
		? [...lines.children].map((c) => c.textContent?.trim()).join(" ")
		: (o.textContent ?? "").trim();
}
const option = (name: string) =>
	[...(surface()?.querySelectorAll<HTMLElement>('[role="option"]') ?? [])].find((o) =>
		o.textContent?.startsWith(name),
	);
const description = (element: Element | undefined) => {
	const id = element?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};
const alertButton = (label: string) =>
	[
		...(document
			.querySelector('[role="alertdialog"]')
			?.querySelectorAll<HTMLButtonElement>("button") ?? []),
	].find((b) => b.textContent?.trim() === label);

async function openPopover() {
	await click(openViewButton());
	await flush(5);
}

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
	stubOverflowWidths({ container: 3000 });
	api.getAllViews.mockResolvedValue({
		systemViews: [active, inactive],
		personalViews: [mine, shared],
	});
	api.whoAmI.mockResolvedValue({ UserId: ME });
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
	api.getAllViews.mockReset();
	api.whoAmI.mockReset();
});

describe("Open view's rules", () => {
	it("waits for an entity, then for its metadata", () => {
		expect(openViewDisabledReason(undefined, undefined)).toBe("Choose an entity first");
		expect(openViewDisabledReason("account", undefined)).toBe("Loading the entity…");
		expect(openViewDisabledReason("account", "accounts")).toBeUndefined();
	});

	it("puts my views first and filters both groups by name, ignoring case", () => {
		expect(openViewGroups([active, inactive], [mine, shared], "")).toEqual({
			mine: [mine, shared],
			system: [active, inactive],
		});
		expect(openViewGroups([active, inactive], [mine, shared], "  ACTIVE ")).toEqual({
			mine: [],
			system: [active, inactive],
		});
		expect(openViewGroups([active, inactive], [mine, shared], "big")).toEqual({
			mine: [mine],
			system: [],
		});
	});
});

describe("Open view on the app bar", () => {
	it("sits between New and Save, and says why it can't open without an entity", async () => {
		await renderBar({ openView: { onOpen: vi.fn() } });
		const names = [...(bar()?.querySelectorAll<HTMLElement>("button") ?? [])]
			.map((b) => b.textContent?.trim())
			.filter(Boolean);
		expect(names.slice(0, 4)).toEqual(["Run", "New", "Open view", "Save"]);
		expect(openViewButton()?.getAttribute("aria-disabled")).toBe("true");
		expect(description(openViewButton())).toBe("Choose an entity first");
		await openPopover();
		expect(surface()).toBeNull();
		expect(api.getAllViews).not.toHaveBeenCalled();
	});

	it("lists my views above system views, marks the default and checks the open view", async () => {
		await renderBar({
			openView: { ...barProps().openView, currentViewId: "s2" },
		});
		expect(openViewButton()?.getAttribute("aria-haspopup")).toBe("dialog");
		await openPopover();
		expect(surface()?.getAttribute("aria-label")).toBe("Open view");
		expect(api.getAllViews).toHaveBeenCalledWith("account");
		expect(groups()).toEqual([
			{ label: "My views", views: ["My big accounts", "Team accounts Shared by Pat Doe"] },
			{
				label: "System views",
				views: ["Active Accounts Default view", "Inactive Accounts"],
			},
		]);
		expect(option("Inactive Accounts")?.getAttribute("aria-selected")).toBe("true");
		expect(option("Active Accounts")?.getAttribute("aria-selected")).toBe("false");
		expect(document.activeElement).toBe(search());
	});

	it("Down in the search box moves into the list", async () => {
		await renderBar();
		await openPopover();
		search()?.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
		await flush(3);
		expect(document.activeElement?.getAttribute("role")).toBe("listbox");
	});

	it("filters as you type, and says when nothing matches", async () => {
		await renderBar();
		await openPopover();
		await typeInto(search(), "team");
		expect(groups()).toEqual([
			{ label: "My views", views: ["Team accounts Shared by Pat Doe"] },
		]);
		await typeInto(search(), "zzz");
		expect(surface()?.textContent).toContain('No views match "zzz".');
	});

	it("opens the chosen view, with what running it by ID needs, and closes", async () => {
		const props = await renderBar();
		await openPopover();
		await click(option("Active Accounts"));
		await flush(5);
		expect(props.openView.onOpen).toHaveBeenCalledWith(
			expect.objectContaining({
				id: "s1",
				type: "system",
				name: "Active Accounts",
				entitySetName: "accounts",
				originalFetchXml: XML,
				layoutxml: "<grid/>",
			}),
		);
		expect(surface()).toBeNull();
	});

	it("asks before it discards unsaved changes: Cancel keeps the query", async () => {
		const props = await renderBar({ viewName: "Mine", unsaved: true });
		await openPopover();
		await click(option("Active Accounts"));
		await flush(5);
		const dialog = document.querySelector('[role="alertdialog"]');
		expect(dialog?.textContent).toContain(
			'Your changes to "Mine" aren\'t saved. Opening "Active Accounts" replaces the tree, the FetchXML and the results.',
		);
		expect(alertButton("Save and open")).toBeTruthy();
		await click(alertButton("Cancel"));
		await flush(5);
		expect(props.openView.onOpen).not.toHaveBeenCalled();
	});

	it("Discard changes opens the view and says what was discarded", async () => {
		const props = await renderBar({ viewName: "Mine", unsaved: true });
		await openPopover();
		await click(option("Active Accounts"));
		await flush(5);
		await click(alertButton("Discard changes"));
		await flush(5);
		expect(props.openView.onOpen).toHaveBeenCalledWith(expect.objectContaining({ id: "s1" }));
		expect(toasts().some((t) => t.includes('Discarded changes to "Mine"'))).toBe(true);
	});

	it("asks from the editor's text too, before the badge has caught up", async () => {
		const props = await renderBar({ unsaved: false, hasUnsavedChanges: () => true });
		await openPopover();
		await click(option("Active Accounts"));
		await flush(5);
		expect(document.querySelector('[role="alertdialog"]')).toBeTruthy();
		expect(props.openView.onOpen).not.toHaveBeenCalled();
	});

	it("loads the views again each time it opens, so a view saved a moment ago is listed", async () => {
		await renderBar();
		await openPopover();
		await click(openViewButton());
		await flush(5);
		expect(surface()).toBeNull();
		api.getAllViews.mockResolvedValue({
			systemViews: [active],
			personalViews: [
				mine,
				view({ id: "p3", name: "Just saved", type: "personal", ownerId: ME }),
			],
		});
		await openPopover();
		expect(api.getAllViews).toHaveBeenCalledTimes(2);
		expect(groups()[0].views).toContain("Just saved");
	});

	it("says when the views can't be loaded, and tries again", async () => {
		vi.spyOn(console, "error").mockImplementation(() => {});
		api.getAllViews.mockRejectedValueOnce(new Error("0x80040220: no access"));
		await renderBar();
		await openPopover();
		expect(surface()?.textContent).toContain("Couldn't load the views. 0x80040220: no access");
		const tryAgain = [...(surface()?.querySelectorAll("button") ?? [])].find(
			(b) => b.textContent?.trim() === "Try again",
		);
		await click(tryAgain);
		await flush(5);
		expect(api.getAllViews).toHaveBeenCalledTimes(2);
		expect(groups()).toHaveLength(2);
	});

	it("says when the entity has no views", async () => {
		api.getAllViews.mockResolvedValue({ systemViews: [], personalViews: [] });
		await renderBar();
		await openPopover();
		expect(surface()?.textContent).toContain("Account has no views you can open.");
	});

	it("opens from More once it has overflowed", async () => {
		const layout = stubOverflowWidths({ container: 3000 });
		await renderBar();
		await layout.resize({ container: 200 });
		const more = bar()?.querySelector<HTMLButtonElement>('button[aria-label="More commands"]');
		await click(more);
		const item = [...document.querySelectorAll<HTMLElement>('[role="menuitem"]')].find(
			(i) => i.textContent?.trim() === "Open view",
		);
		expect(item?.getAttribute("aria-haspopup")).toBe("dialog");
		await click(item);
		await new Promise((resolve) => setTimeout(resolve, 0));
		await flush(5);
		expect(surface()).toBeTruthy();
		expect(groups()).toHaveLength(2);
	});
});

describe("Open view after a failed load", () => {
	it("doesn't say it discarded anything when the view couldn't be opened", async () => {
		const onOpen = vi.fn(() => false);
		await renderBar({
			viewName: "Mine",
			unsaved: true,
			openView: { ...barProps().openView, onOpen },
		});
		await openPopover();
		await click(option("Active Accounts"));
		await flush(5);
		await click(alertButton("Discard changes"));
		await flush(5);
		expect(onOpen).toHaveBeenCalled();
		expect(toasts().some((t) => t.includes("Discarded"))).toBe(false);
	});
});

describe("a popover command that becomes unavailable while open", () => {
	it("tells its owner it closed, so it doesn't reopen by itself", async () => {
		const onPopoverOpenChange = vi.fn();
		let setReason: (reason: string | undefined) => void = () => {};
		function Bar() {
			const [reason, set] = useState<string | undefined>(undefined);
			setReason = set;
			return createElement(OverflowToolbar, {
				ariaLabel: "Test",
				commands: [
					{
						id: "pick",
						label: "Pick",
						kind: "popover",
						popover: createElement("p", null, "Inside"),
						popoverOpen: true,
						onPopoverOpenChange,
						disabledReason: reason,
					},
				],
			});
		}
		await render(createElement(Bar));
		await flush(5);
		expect(surface()?.textContent).toBe("Inside");
		expect(onPopoverOpenChange).not.toHaveBeenCalled();
		await act(async () => setReason("Choose an entity first"));
		await flush(5);
		expect(onPopoverOpenChange).toHaveBeenCalledWith(false);
		expect(surface()).toBeNull();
	});
});
