/**
 * The right pane's tab row (#232, #278): tabs never wrap, the selected tab stays visible, and
 * choosing a tab from More selects it. The commands are in the app bar (AppBar.test.ts). jsdom has
 * no layout, so widths come from stubOverflowWidths.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../../../shared/testing/ui";
import { PreviewHeader, type PreviewHeaderProps } from "./PreviewHeader";

const renderHeader = (props: Partial<PreviewHeaderProps> = {}) =>
	render(
		createElement(PreviewHeader, {
			selectedTab: "xml",
			onTabSelect: () => {},
			advisoryCount: 2,
			...props,
		}),
	);

const tabList = () => document.querySelector<HTMLElement>('[role="tablist"]');
const overflowed = (element: Element | null | undefined) =>
	!!element?.closest("[data-overflowing]");
const tab = (label: string) =>
	[...(tabList()?.querySelectorAll<HTMLElement>('[role="tab"]') ?? [])].find(
		(t) => t.getAttribute("aria-label") === label || t.textContent?.startsWith(label),
	);
const moreTabs = () =>
	tabList()?.querySelector<HTMLButtonElement>('button[aria-label^="More tabs"]');
const menus = () => [...document.querySelectorAll<HTMLElement>('[role="menu"]')];

beforeEach(() => {
	// Fluent's positioning warns in jsdom
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("the header has only tabs", () => {
	it("no commands: they're in the app bar", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader();
		expect(document.querySelector('[role="toolbar"]')).toBeNull();
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
