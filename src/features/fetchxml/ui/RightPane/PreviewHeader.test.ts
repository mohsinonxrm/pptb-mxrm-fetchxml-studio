/**
 * The right pane's tab row (#232, #278, #233): tabs never wrap, the selected tab stays visible, and
 * choosing a tab from More selects it. Every tab has an icon; Results and Advisories show counts, in
 * their accessible names too, and keep them in More. The commands are in the app bar (AppBar.test.ts). jsdom has
 * no layout, so widths come from stubOverflowWidths.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../../../shared/testing/ui";
import { PreviewHeader, type PreviewHeaderProps } from "./PreviewHeader";
import type { Advisory, AdvisoryLevel } from "../../model/advisories";

const advisory = (level: AdvisoryLevel, id: string) =>
	({ id, level, title: id }) as unknown as Advisory;

const renderHeader = (props: Partial<PreviewHeaderProps> = {}) =>
	render(
		createElement(PreviewHeader, {
			selectedTab: "xml",
			onTabSelect: () => {},
			advisories: [advisory("advice", "a1"), advisory("advice", "a2")],
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

const badgeOf = (element: Element | undefined) => element?.querySelector<HTMLElement>(".fui-Badge");

describe("the header's tabs", () => {
	it("each have an icon and their text", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader();
		const tabs = [...(tabList()?.querySelectorAll<HTMLElement>('[role="tab"]') ?? [])];
		const labels = ["FetchXML", "LayoutXML", "Results", "Advisories", "Code"];
		expect(tabs).toHaveLength(labels.length);
		tabs.forEach((t, i) =>
			expect(t.querySelector(".fui-Tab__content")?.textContent?.startsWith(labels[i])).toBe(
				true,
			),
		);
		for (const t of tabs) expect(t.querySelector(".fui-Tab__icon svg")).toBeTruthy();
		expect(tabList()?.getAttribute("aria-label")).toBe("Query views");
	});

	it("show the results' row count in the user's locale, in the name too; none before a run", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader();
		expect(badgeOf(tab("Results"))).toBeNull();
		expect(tab("Results")?.getAttribute("aria-label")).toBe("Results");
		cleanup();
		await renderHeader({ rowCount: 48213 });
		const results = tab("Results, 48,213 rows");
		expect(badgeOf(results)?.textContent).toBe(new Intl.NumberFormat().format(48213));
		cleanup();
		await renderHeader({ rowCount: 1 });
		expect(tab("Results, 1 row")).toBeTruthy();
	});

	it("show the advisories' total, named by level, coloured and marked by the worst one", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({
			advisories: [
				advisory("warning", "w1"),
				advisory("error", "e1"),
				advisory("advice", "a1"),
				advisory("warning", "w2"),
			],
		});
		const tabEl = tab(
			"Advisories, 4: 1 will be rejected, 2 might be rejected, 1 performance advice",
		);
		const badge = badgeOf(tabEl);
		expect(badge?.textContent).toBe("4");
		expect(badge?.dataset.color).toBe("danger");
		// Not by colour alone: errors have their own icon
		expect(badge?.querySelector("svg")).toBeTruthy();
	});

	it("colour the advisories by warning or advice when there's no error, and hide the badge at 0", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderHeader({ advisories: [advisory("warning", "w1")] });
		expect(badgeOf(tab("Advisories, 1: 1 might be rejected"))?.dataset.color).toBe("warning");
		cleanup();
		await renderHeader({ advisories: [advisory("advice", "a1")] });
		const advice = badgeOf(tab("Advisories, 1: 1 performance advice"));
		expect(advice?.dataset.color).toBe("informative");
		expect(advice?.querySelector("svg")).toBeNull();
		cleanup();
		await renderHeader({ advisories: [] });
		expect(badgeOf(tab("Advisories"))).toBeNull();
	});

	it("keep the count and the full name in More", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		await renderHeader({ selectedTab: "xml", rowCount: 25 });
		await layout.resize({ container: 150 });
		await click(moreTabs());
		const item = [
			...(menus()[0]?.querySelectorAll<HTMLElement>('[role="menuitem"]') ?? []),
		].find((i) => i.getAttribute("aria-label") === "Results, 25 rows");
		expect(badgeOf(item)?.textContent).toBe("25");
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
