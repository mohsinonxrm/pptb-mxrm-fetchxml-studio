/**
 * Headings and names (#279 PR 2, #390): the Properties pane is an h2 and the editors' sections are
 * h3; the Settings drawer's sections are h3 and its switch and dropdown are named by their labels;
 * the scope popover is named by its h2 title; the scope chip isn't a second primary button.
 */
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { cleanup, click, removeHost, stubOverflowWidths } from "../shared/testing/ui";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = () => React.createElement("textarea", { "data-monaco": "", readOnly: true });
	return {
		default: Stub,
		Editor: Stub,
		DiffEditor: Stub,
		loader: { config() {}, init: async () => ({}) },
		useMonaco: () => null,
	};
});

afterEach(() => {
	cleanup();
	removeHost();
});

const FLOW_TIMEOUT = 120000;
const QUERY =
	"<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/></entity></fetch>";
type App = Awaited<ReturnType<typeof renderApp>>;
const settle = (ms = 400) => act(() => new Promise<void>((resolve) => setTimeout(resolve, ms)));
const named = (name: string | RegExp) =>
	[...document.querySelectorAll<HTMLElement>("button")].find((b) => {
		const n = b.getAttribute("aria-label") ?? b.textContent?.trim() ?? "";
		return typeof name === "string" ? n === name : name.test(n);
	});
const texts = (selector: string, root: ParentNode = document) =>
	[...root.querySelectorAll(selector)].map((h) => h.textContent?.trim());
const labelOf = (el: Element | null | undefined) =>
	(el?.getAttribute("aria-labelledby") ?? "")
		.split(/\s+/)
		.map((id) => document.getElementById(id)?.textContent?.trim() ?? "")
		.join(" ")
		.trim();
const ready = async (app: App) => {
	await app.waitFor(() => !!document.querySelector('[role="treeitem"]'));
	const gotIt = named("Got it");
	if (gotIt) await click(gotIt);
	await app.waitFor(() => !document.querySelector('[role="dialog"][aria-modal="true"]'));
	await settle();
};

describe("app: headings and names", () => {
	it(
		"Properties is an h2, and the selected node's sections are h3",
		async () => {
			stubOverflowWidths({ container: 4000 });
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			expect(texts("h2")).toContain("Properties");
			const fetchRow = [...document.querySelectorAll<HTMLElement>('[role="treeitem"]')].find(
				(r) =>
					r.querySelector(".fui-TreeItemLayout__main")?.textContent?.trim() ===
					"Fetch query",
			);
			await click(fetchRow?.querySelector(".fui-TreeItemLayout"));
			await settle(200);
			// The fetch options editor's sections
			await app.waitFor(() => texts("h3").includes("Run options"));
			expect(texts("h3")).toEqual(
				expect.arrayContaining(["Run options", "Query", "Paging and limits"]),
			);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"Settings: h3 sections, and the switch and the dropdown named by their labels (#390)",
		async () => {
			stubOverflowWidths({ container: 4000 });
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await click(named("Settings"));
			const settingsDrawer = () =>
				[...document.querySelectorAll('[role="dialog"]')].find((d) =>
					d.textContent?.includes("Value display mode"),
				);
			await app.waitFor(() => !!settingsDrawer());
			const drawer = settingsDrawer()!;
			expect(texts("h3", drawer)).toEqual(expect.arrayContaining(["Query scope", "Display"]));
			expect(labelOf(drawer.querySelector('[role="switch"]'))).toBe("Use logical names");
			const dropdown = [...drawer.querySelectorAll('[role="combobox"]')].find(
				(c) => labelOf(c) === "Value display mode",
			);
			expect(dropdown).toBeTruthy();
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"the scope chip is outline, and its popover is named by its h2 title",
		async () => {
			stubOverflowWidths({ container: 4000 });
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			const chip = named(/^(Choose scope|Entity scope)/);
			// Run is the only button with the brand background (Griffel styles resolve in jsdom)
			const brand = [...document.querySelectorAll<HTMLElement>("button")].filter(
				(b) => getComputedStyle(b).backgroundColor === "var(--colorBrandBackground)",
			);
			expect(brand.map((b) => b.textContent?.trim())).toEqual(["Run"]);
			expect(brand).not.toContain(chip);
			await click(chip);
			const surface = () =>
				[...document.querySelectorAll('[role="dialog"]')].find((d) =>
					d.textContent?.includes("Entity scope"),
				);
			await app.waitFor(() => !!surface());
			expect(labelOf(surface())).toBe("Entity scope");
			expect(texts("h2", surface())).toContain("Entity scope");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});

describe("app: the column filter popover (#279)", () => {
	it(
		"is named by its h2 title",
		async () => {
			stubOverflowWidths({ container: 4000 });
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await click(named("Run"));
			await app.waitFor(() => /Rows: \d/.test(app.status()));
			const header = [
				...document.querySelectorAll<HTMLElement>('[role="columnheader"]'),
			].find((h) => /Account Name|name/.test(h.textContent ?? ""));
			await click(header?.querySelector("button") ?? header);
			const filterItem = () =>
				[...document.querySelectorAll<HTMLElement>('[role="menuitem"]')].find((i) =>
					i.textContent?.startsWith("Filter by"),
				);
			await app.waitFor(() => !!filterItem());
			await click(filterItem());
			const surface = () =>
				[...document.querySelectorAll('[role="dialog"]')].find((d) =>
					d.textContent?.includes("Filter by"),
				);
			await app.waitFor(() => !!surface());
			expect(labelOf(surface())).toMatch(/^Filter by (Account Name|name)$/);
			expect(texts("h2", surface())).toEqual([labelOf(surface())]);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});
