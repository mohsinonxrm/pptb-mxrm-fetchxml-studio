/**
 * The right pane's header row on the shared overflow bars (#232): tabs and commands never wrap,
 * Execute and Settings never overflow, the selected tab stays visible, disabled commands keep their
 * reason, and Save's menu shows once in More (no second Save). jsdom has no layout, so widths come
 * from stubOverflowWidths.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../../../shared/testing/ui";
import { PreviewHeader, type PreviewHeaderProps } from "./PreviewHeader";
import { EMPTY_EDITOR_REASON, emptyQueryReason } from "./previewHeaderRules";
import { NO_QUERY_REASON } from "../Toolbar/useSaveViewCommand";
import { EditorSyncContext } from "./editorSync";

const execute = (extra: Partial<PreviewHeaderProps["execute"]> = {}) => ({
	onClick: vi.fn(),
	isExecuting: false,
	busy: false,
	...extra,
});

const saveView = (extra: Partial<NonNullable<PreviewHeaderProps["saveView"]>> = {}) => ({
	fetchXml: "<fetch><entity name='account'/></fetch>",
	layoutXml: "",
	entityLogicalName: "account",
	objectTypeCode: 1,
	primaryIdAttribute: "accountid",
	loadedView: null,
	...extra,
});

const renderHeader = (props: Partial<PreviewHeaderProps> = {}) =>
	render(
		createElement(PreviewHeader, {
			selectedTab: "xml",
			onTabSelect: () => {},
			advisoryCount: 2,
			execute: execute(),
			saveView: saveView(),
			getCurrentXml: () => "",
			...props,
		}),
	);

const commandBar = () => document.querySelector<HTMLElement>('[role="toolbar"]');
const tabList = () => document.querySelector<HTMLElement>('[role="tablist"]');
const overflowed = (element: Element | null | undefined) =>
	!!element?.closest("[data-overflowing]");
/** A bar control by its accessible name or text (a split's primary action for a split) */
const onBar = (name: string) =>
	[...(commandBar()?.querySelectorAll<HTMLElement>("button") ?? [])].find(
		(b) => b.getAttribute("aria-label") === name || b.textContent?.trim() === name,
	);
const shownOnBar = (names: string[]) => names.filter((n) => onBar(n) && !overflowed(onBar(n)));
const tab = (label: string) =>
	[...(tabList()?.querySelectorAll<HTMLElement>('[role="tab"]') ?? [])].find(
		(t) => t.getAttribute("aria-label") === label || t.textContent?.startsWith(label),
	);
const moreCommands = () =>
	commandBar()?.querySelector<HTMLButtonElement>('button[aria-label="More commands"]');
const moreTabs = () =>
	tabList()?.querySelector<HTMLButtonElement>('button[aria-label^="More tabs"]');
const menus = () => [...document.querySelectorAll<HTMLElement>('[role="menu"]')];
const menuTexts = (menu: HTMLElement | undefined) =>
	[...(menu?.querySelectorAll<HTMLElement>('[role^="menuitem"]') ?? [])].map((i) =>
		i.textContent?.trim(),
	);
const description = (element: Element | undefined) => {
	const id = element?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};

beforeEach(() => {
	// Fluent's positioning warns in jsdom
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("the header's commands", () => {
	it("are Execute, Return FetchXML, Save and Send to tool, then Settings", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({
			isCallee: true,
			sendToTool: {
				fetchXml: "<fetch/>",
				entityLogicalName: "account",
				tools: [{ id: "other", name: "Other" }],
			},
		});
		const names = [...(commandBar()?.querySelectorAll<HTMLElement>("button") ?? [])]
			.map((b) => b.getAttribute("aria-label") ?? b.textContent?.trim())
			.filter(Boolean);
		expect(names).toEqual([
			"Execute",
			"Return FetchXML",
			// The split's menu half has neither text nor a name of its own
			"Save as",
			"Send to tool",
			"Settings",
		]);
		expect(commandBar()?.getAttribute("aria-label")).toBe("Query commands");
	});

	it("shows Save on the FetchXML tab only", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({ selectedTab: "results" });
		expect(onBar("Save as")).toBeUndefined();
		await renderHeader({ selectedTab: "xml" });
		expect(onBar("Save as")).toBeTruthy();
	});

	it("keeps Execute and Settings when narrow, and moves the rest into More in bar order", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		await renderHeader({
			isCallee: true,
			sendToTool: {
				fetchXml: "<fetch/>",
				entityLogicalName: "account",
				tools: [{ id: "other", name: "Other" }],
			},
		});
		await layout.resize({ container: 200 });
		expect(
			shownOnBar(["Execute", "Return FetchXML", "Save as", "Send to tool", "Settings"]),
		).toEqual(["Execute", "Settings"]);
		await click(moreCommands());
		expect(menuTexts(menus()[0])).toEqual(["Return FetchXML", "Save as", "Send to tool"]);
	});

	it("never shows two dividers in a row, or one with nothing after it, as commands overflow", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		await renderHeader({
			isCallee: true,
			sendToTool: {
				fetchXml: "<fetch/>",
				entityLogicalName: "account",
				tools: [{ id: "other", name: "Other" }],
			},
		});
		// What shows on the bar, in order: "|" for a divider, a name for a control
		const sequence = () =>
			[...(commandBar()?.children ?? [])]
				.filter(
					(c) =>
						!c.hasAttribute("data-overflowing") &&
						getComputedStyle(c).display !== "none",
				)
				.map((c) => {
					if (
						c.getAttribute("role") === "separator" ||
						c.querySelector('[role="separator"]')
					)
						return "|";
					const b = c.matches("button") ? c : c.querySelector("button");
					return b?.getAttribute("aria-label") ?? b?.textContent?.trim() ?? "";
				})
				.filter(Boolean);
		for (const width of [700, 600, 520, 480, 440, 400, 360, 300]) {
			await layout.resize({ container: width });
			const shown = sequence();
			const doubled = shown.some((s, i) => s === "|" && shown[i + 1] === "|");
			expect({ width, shown, doubled }).toEqual({ width, shown, doubled: false });
			expect(shown[0]).not.toBe("|");
		}
	});

	it("shows Save's own menu in More, without a second Save", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		await renderHeader({
			saveView: saveView({
				loadedView: { id: "1", name: "Active accounts", type: "system" },
			}),
		});
		await layout.resize({ container: 200 });
		await click(moreCommands());
		const save = [
			...(menus()[0]?.querySelectorAll<HTMLElement>('[role="menuitem"]') ?? []),
		].find((i) => i.textContent?.trim() === "Save");
		await click(save);
		expect(menuTexts(menus()[1])).toEqual(["Save", "Save as…", "Save and publish"]);
	});

	it("keeps an unavailable Execute and Save focusable, with the reason", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({
			execute: execute({ disabledReason: NO_QUERY_REASON }),
			saveView: saveView({ disabled: true }),
		});
		const run = onBar("Execute");
		expect(run?.getAttribute("aria-disabled")).toBe("true");
		expect(description(run)).toBe("Choose an entity first");
		const save = onBar("Save as");
		expect(save?.getAttribute("aria-disabled")).toBe("true");
		expect(description(save)).toBe("Choose an entity first");
	});

	it("says Executing… while a query runs, and can't run another", async () => {
		stubOverflowWidths({ container: 2000 });
		const onClick = vi.fn();
		await renderHeader({ execute: execute({ isExecuting: true, busy: true, onClick }) });
		const run = onBar("Executing…");
		expect(run?.getAttribute("aria-disabled")).toBe("true");
		await click(run);
		expect(onClick).not.toHaveBeenCalled();
	});

	it("names Settings, which is icon only", async () => {
		stubOverflowWidths({ container: 2000 });
		const onOpenSettings = vi.fn();
		await renderHeader({ onOpenSettings });
		const settings = onBar("Settings");
		expect(settings?.textContent?.trim()).toBe("");
		await click(settings);
		expect(onOpenSettings).toHaveBeenCalledOnce();
	});
});

describe("which commands show", () => {
	const tools = [{ id: "other", name: "Other" }];

	it("shows Return FetchXML only when another tool opened FetchXML Studio", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader();
		expect(onBar("Return FetchXML")).toBeUndefined();
		await renderHeader({ isCallee: true });
		expect(onBar("Return FetchXML")).toBeTruthy();
	});

	it("shows Send to tool only when a tool takes FetchXML", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({
			sendToTool: { fetchXml: "<fetch/>", entityLogicalName: "account", tools: [] },
		});
		expect(onBar("Send to tool")).toBeUndefined();
		await renderHeader({
			sendToTool: { fetchXml: "<fetch/>", entityLogicalName: "account", tools },
		});
		expect(onBar("Send to tool")?.getAttribute("aria-disabled")).toBeNull();
	});

	it("keeps Send to tool focusable with its reason before there's a query", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({
			sendToTool: { fetchXml: "", entityLogicalName: "", tools, disabled: true },
		});
		const send = onBar("Send to tool");
		expect(send?.getAttribute("aria-disabled")).toBe("true");
		expect(description(send)).toBe("Choose an entity first");
	});

	it("offers Save and Send to tool with an empty tree when the editor has FetchXML", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(
			createElement(
				EditorSyncContext.Provider,
				{
					value: {
						sync: async () => ({ status: "unchanged" as const }),
						editorHasContent: true,
					},
				},
				createElement(PreviewHeader, {
					selectedTab: "xml",
					onTabSelect: () => {},
					advisoryCount: 0,
					execute: execute(),
					saveView: saveView({ disabled: true }),
					sendToTool: { fetchXml: "", entityLogicalName: "", tools, disabled: true },
					getCurrentXml: () => "",
				}),
			),
		);
		expect(onBar("Save as")?.getAttribute("aria-disabled")).toBeNull();
		expect(onBar("Send to tool")?.getAttribute("aria-disabled")).toBeNull();
	});
});

describe("Execute's reason for an empty query", () => {
	it("asks for an entity, or for FetchXML in editor mode", () => {
		expect(emptyQueryReason(false, false, "")).toBe(NO_QUERY_REASON);
		expect(emptyQueryReason(false, false, "<fetch/>")).toBeUndefined();
		expect(emptyQueryReason(true, false, "<fetch/>")).toBe(EMPTY_EDITOR_REASON);
		expect(emptyQueryReason(true, true, "")).toBeUndefined();
	});
});

describe("the header's tabs", () => {
	it("name Advisories with its count", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({ advisoryCount: 3 });
		expect(tab("Advisories (3)")).toBeTruthy();
		expect(tabList()?.getAttribute("aria-label")).toBe("Query views");
	});

	it("keep the selected tab when narrow, and choosing one from More selects it", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		const onTabSelect = vi.fn();
		await renderHeader({ selectedTab: "code", onTabSelect });
		// Room for one tab and More
		await layout.resize({ container: 150 });
		expect(overflowed(tab("Code"))).toBe(false);
		expect(overflowed(tab("FetchXML"))).toBe(true);
		await click(moreTabs());
		const fetchXml = [
			...(menus()[0]?.querySelectorAll<HTMLElement>('[role="menuitem"]') ?? []),
		].find((i) => i.textContent?.trim() === "FetchXML");
		await click(fetchXml);
		expect(onTabSelect).toHaveBeenCalledWith("xml");
	});
});
