/**
 * The shared overflow tab row (#232): tabs that don't fit move into "More", the selected tab always
 * stays visible, and choosing a tab from More selects it. Widths come from stubOverflowWidths (ui.ts).
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../testing/ui";
import { OverflowTabList, type OverflowTabListProps } from "./OverflowTabList";
import type { BarTab } from "./types";

// Every tab is 80 px and More 40 px; the tab list keeps Fluent's default 10 px free
const PADDING = 10;
const fits = (px: number) => PADDING + px;

const tabs: BarTab[] = ["FetchXML", "LayoutXML", "Results", "Advisories", "Code"].map((label) => ({ value: label.toLowerCase(), label }));
const labels = tabs.map((t) => t.label);

const renderTabs = (selectedValue: string, onTabSelect: OverflowTabListProps["onTabSelect"] = () => {}) =>
	render(createElement(OverflowTabList, { ariaLabel: "Query views", tabs, selectedValue, onTabSelect }));

const tabList = () => document.querySelector<HTMLElement>('[role="tablist"]');
// A Tab repeats its text in a hidden span that reserves the selected (bold) width, so read the content slot
const tab = (label: string) => [...(tabList()?.querySelectorAll<HTMLElement>('[role="tab"]') ?? [])].find((t) => t.querySelector(".fui-Tab__content")?.textContent?.trim() === label);
const overflowed = (element: Element | undefined | null) => !!element?.closest("[data-overflowing]");
const visibleTabs = () => labels.filter((l) => !overflowed(tab(l)));
const more = () => tabList()?.querySelector<HTMLButtonElement>('button[aria-label^="More tabs"]');
const menuItems = () => [...document.querySelectorAll<HTMLElement>('[role="menu"] [role="menuitem"]')];

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("OverflowTabList", () => {
	it("shows every tab and no More when they fit", async () => {
		stubOverflowWidths({ container: fits(5 * 80) });
		await renderTabs("fetchxml");
		expect(visibleTabs()).toEqual(labels);
		expect(more()).toBeFalsy();
	});

	it("keeps the selected tab visible when the row narrows", async () => {
		const layout = stubOverflowWidths({ container: 1000 });
		await renderTabs("code");
		// Room for 2 tabs and More: the selected last tab stays, with the first one
		await layout.resize({ container: fits(2 * 80 + 40) });
		expect(visibleTabs()).toEqual(["FetchXML", "Code"]);
		expect(tab("Code")?.getAttribute("aria-selected")).toBe("true");

		// Only room for one: it's the selected one
		await layout.resize({ container: fits(80 + 40) });
		expect(visibleTabs()).toEqual(["Code"]);
	});

	it("labels More with text, the hidden count in its name, and lists the hidden tabs in order", async () => {
		stubOverflowWidths({ container: fits(2 * 80 + 40) });
		await renderTabs("fetchxml");
		const moreButton = more();
		expect(moreButton?.textContent?.trim()).toBe("More");
		expect(moreButton?.getAttribute("aria-label")).toBe("More tabs (3)");
		expect(moreButton?.getAttribute("role")).toBe("tab");

		await click(moreButton);
		expect(menuItems().map((i) => i.textContent?.trim())).toEqual(["Results", "Advisories", "Code"]);
	});

	it("selects a tab chosen from More, which then shows on the row", async () => {
		stubOverflowWidths({ container: fits(2 * 80 + 40) });
		const onTabSelect = vi.fn();
		await renderTabs("fetchxml", onTabSelect);
		expect(overflowed(tab("Advisories"))).toBe(true);

		await click(more());
		await click(menuItems().find((i) => i.textContent?.trim() === "Advisories"));
		expect(onTabSelect).toHaveBeenCalledWith("advisories");

		// The owner selects it: it outranks the others now, so it's back on the row
		await renderTabs("advisories", onTabSelect);
		expect(overflowed(tab("Advisories"))).toBe(false);
		expect(tab("Advisories")?.getAttribute("aria-selected")).toBe("true");
		expect(visibleTabs()).toHaveLength(2);
	});

	it("calls onTabSelect when a visible tab is clicked", async () => {
		stubOverflowWidths({ container: 1000 });
		const onTabSelect = vi.fn();
		await renderTabs("fetchxml", onTabSelect);
		await click(tab("Results"));
		expect(onTabSelect).toHaveBeenCalledWith("results");
	});
});
