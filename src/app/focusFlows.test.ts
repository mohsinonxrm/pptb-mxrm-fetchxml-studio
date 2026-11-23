/**
 * Focus, tree and grid keys in the whole tool (#186 PR 2):
 *  - Delete on a focused tree row deletes it (a node with items inside asks first), and focus goes
 *    to the next row at the same level, else the parent; not with a modifier, not while typing.
 *  - Closing a dialog puts focus back on what opened it (the row whose menu opened it).
 *  - Shift+F10 or the menu key on a focused results cell opens the cell menu.
 *  - The sorted column header says so (aria-sort), and the grid is named "Results: {entity}".
 */
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, openDialog, removeHost } from "../shared/testing/ui";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = (props: { value?: string; defaultValue?: string }) =>
		React.createElement("textarea", {
			"data-monaco": "",
			readOnly: true,
			value: props.value ?? props.defaultValue ?? "",
		});
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
	"<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/><filter><condition attribute='name' operator='not-null'/></filter></entity></fetch>";

const rowLabel = (el: Element | null) =>
	el?.querySelector(".fui-TreeItemLayout__main")?.textContent?.trim();
/** A tree row by its exact label ("name", not the sort row "name ↑"), or its start with a "*" suffix */
const treeItem = (text: string) =>
	[...document.querySelectorAll<HTMLElement>('[role="treeitem"]')].find((r) => {
		const label = rowLabel(r) ?? "";
		return text.endsWith("*") ? label.startsWith(text.slice(0, -1)) : label === text;
	});
const press = (target: Element | null | undefined, init: KeyboardEventInit) =>
	act(async () => {
		target?.dispatchEvent(
			new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
		);
	});
const active = () => document.activeElement as HTMLElement | null;
/** The first-run scope tip traps focus: dismiss it as a user would */
const ready = async (app: Awaited<ReturnType<typeof renderApp>>) => {
	await app.waitFor(() => !!treeItem("name"));
	const gotIt = [...document.querySelectorAll<HTMLElement>("button")].find(
		(b) => b.textContent?.trim() === "Got it",
	);
	if (gotIt) await click(gotIt);
	await app.waitFor(() => !document.querySelector('[role="dialog"][aria-modal="true"]'));
	// Let the tip's close settle, as it would before a person's next key press (a popup right
	// after another counts as the same run: useFocusReturn)
	await act(() => new Promise<void>((resolve) => setTimeout(resolve, 400)));
};

describe("app: Delete on a tree row", () => {
	it(
		"deletes the focused row, and focus goes to the next row at the same level",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await act(async () => treeItem("name")?.focus());
			await press(active(), { key: "Delete" });
			await app.waitFor(() => !treeItem("name"));
			await app.waitFor(() => rowLabel(active()) === "accountid");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"with a modifier, or while typing in a field, Delete deletes nothing",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await act(async () => treeItem("name")?.focus());
			await press(active(), { key: "Delete", ctrlKey: true });
			await press(active(), { key: "Delete", shiftKey: true });
			expect(treeItem("name")).toBeTruthy();
			const field = document.querySelector<HTMLInputElement>(
				'aside input[type="text"], aside input:not([type])',
			);
			expect(field).toBeTruthy();
			await press(field, { key: "Delete" });
			// Held down, it deletes once, not row after row
			await act(async () => treeItem("name")?.focus());
			await press(active(), { key: "Delete", repeat: true });
			expect(treeItem("name")).toBeTruthy();
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"a node with items inside asks first; Cancel puts focus back on the row, Delete moves it on",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			const filter = () => treeItem("Filter*");
			await act(async () => filter()?.focus());
			await press(active(), { key: "Delete" });
			await app.waitFor(() => !!openDialog());
			await click(button("Cancel"));
			await app.waitFor(() => !openDialog());
			// Back on the row that asked
			await app.waitFor(() => active() === filter());

			await press(active(), { key: "Delete" });
			await app.waitFor(() => !!openDialog());
			await click(button("Delete"));
			await app.waitFor(() => !filter());
			// The filter was the last row of the entity: focus goes to the entity
			await app.waitFor(() => rowLabel(active())?.startsWith("account") === true);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});

describe("app: the results grid", () => {
	const runQuery = async (app: Awaited<ReturnType<typeof renderApp>>) => {
		await click(button("Run", document));
		await app.waitFor(() => /Rows: \d/.test(app.status()));
	};
	const cells = () => [...document.querySelectorAll<HTMLElement>('[role="gridcell"]')];

	it(
		"Shift+F10 or the menu key on a focused cell opens the cell menu",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await runQuery(app);
			const cell = cells().find((c) => c.textContent?.trim());
			await act(async () => cell?.focus());
			await press(cell, { key: "F10", shiftKey: true });
			const menu = () => document.querySelector<HTMLElement>('[role="menu"]');
			await app.waitFor(() => !!menu());
			expect(menu()?.textContent).toContain("Copy");
			await press(menu(), { key: "Escape" });
			await app.waitFor(() => !menu());
			await press(cell, { key: "ContextMenu" });
			await app.waitFor(() => !!menu());
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"with two orders, only the first one's header is announced as sorted",
		async () => {
			const twoOrders = QUERY.replace(
				"<order attribute='name'/>",
				"<order attribute='name'/><order attribute='accountid' descending='true'/>",
			);
			const app = await renderApp({ launch: { fetchXml: twoOrders } });
			await ready(app);
			await runQuery(app);
			const sorted = [...document.querySelectorAll<HTMLElement>('[role="columnheader"]')]
				.map((h) => h.getAttribute("aria-sort"))
				.filter((sort) => sort === "ascending" || sort === "descending");
			expect(sorted).toEqual(["ascending"]);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"names the grid for its entity, and the sorted column header says so",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await runQuery(app);
			const grid = document.querySelector('[role="grid"]');
			expect(grid?.getAttribute("aria-label")).toBe("Results: account");
			const headers = [...document.querySelectorAll<HTMLElement>('[role="columnheader"]')];
			const sorts = headers
				.filter((h) => h.hasAttribute("aria-sort"))
				.map((h) => [h.textContent?.trim(), h.getAttribute("aria-sort")]);
			expect(sorts.find(([, sort]) => sort === "ascending")?.[0]).toMatch(
				/Account Name|name/,
			);
			expect(sorts.filter(([, sort]) => sort === "ascending")).toHaveLength(1);
			expect(sorts.some(([, sort]) => sort === "none")).toBe(true);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});
