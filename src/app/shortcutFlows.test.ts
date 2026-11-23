/**
 * The tool's keys in the whole app (#186): Ctrl+S opens Save as selecting Save does, Ctrl+/ and
 * More › Keyboard shortcuts open the list, Ctrl+F6 / Ctrl+Shift+F6 move between the app bar, the
 * builder and the output, and F2 goes to the newest toast. The app bar is a banner landmark, and
 * Save shows its key.
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
	"<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/></entity></fetch>";

const treeItem = (text: string) =>
	[...document.querySelectorAll<HTMLElement>('[role="treeitem"]')].find((r) =>
		r.querySelector(".fui-TreeItemLayout__main")?.textContent?.trim().startsWith(text),
	);
const press = (target: Element | null | undefined, init: KeyboardEventInit) =>
	act(async () => {
		target?.dispatchEvent(
			new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
		);
	});
const header = () => document.querySelector<HTMLElement>("header");
const output = () => document.querySelector<HTMLElement>('main[aria-label="Query output"]');
const pane = () => document.querySelector<HTMLElement>('aside[aria-label="Query builder"]');
const active = () => document.activeElement as HTMLElement | null;
const dialogTitle = () => openDialog()?.querySelector("h2")?.textContent;
/**
 * The first-run scope tip traps focus (a modal dialog in front), so the keys wait for it, as they
 * should: dismiss it the way a user would
 */
const ready = async (app: Awaited<ReturnType<typeof renderApp>>) => {
	await app.waitFor(() => !!treeItem("name"));
	const gotIt = [...document.querySelectorAll<HTMLElement>("button")].find(
		(b) => b.textContent?.trim() === "Got it",
	);
	if (gotIt) await click(gotIt);
	await app.waitFor(() => !document.querySelector('[role="dialog"][aria-modal="true"]'));
};

describe("app: keyboard shortcuts", () => {
	it(
		"Ctrl+S opens Save, and Save shows its key",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			const save = button("Save", document);
			expect(save?.getAttribute("aria-keyshortcuts")).toBe("Control+S");
			await press(treeItem("Fetch query"), { key: "s", ctrlKey: true });
			await app.waitFor(() => /Save/.test(dialogTitle() ?? ""));
			expect(dialogTitle()).toBe("Save as new view");
			// Not expectClean: the Save dialog's performance check isn't in the fake Dataverse
		},
		FLOW_TIMEOUT,
	);

	it(
		"Ctrl+/ and More › Keyboard shortcuts open the list",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await press(treeItem("Fetch query"), { key: "/", ctrlKey: true });
			await app.waitFor(() => dialogTitle() === "Keyboard shortcuts");
			expect(openDialog()?.textContent).toContain("Run query");
			await click(button("Close"));
			await app.waitFor(() => !openDialog());

			await click(
				document.querySelector('header button[aria-label="More commands"]') as HTMLElement,
			);
			const item = [...document.querySelectorAll<HTMLElement>('[role="menuitem"]')].find(
				(i) => i.textContent?.startsWith("Keyboard shortcuts"),
			);
			expect(item?.textContent).toBe("Keyboard shortcutsCtrl+/");
			await click(item);
			await app.waitFor(() => dialogTitle() === "Keyboard shortcuts");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"Ctrl+F6 and Ctrl+Shift+F6 move between the app bar, the builder and the output",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			expect(
				header()?.querySelector('[role="toolbar"][aria-label="Query commands"]'),
			).toBeTruthy();
			await act(async () => treeItem("name")?.focus());
			await press(active(), { key: "F6", ctrlKey: true });
			// The output's first stop is its selected tab
			expect(output()?.contains(active())).toBe(true);
			expect(active()?.getAttribute("role")).toBe("tab");
			await press(active(), { key: "F6", ctrlKey: true });
			expect(header()?.contains(active())).toBe(true);
			await press(active(), { key: "F6", ctrlKey: true });
			// Back to where focus was in the builder
			expect(pane()?.contains(active())).toBe(true);
			expect(active()).toBe(treeItem("name"));
			await press(active(), { key: "F6", ctrlKey: true, shiftKey: true });
			expect(header()?.contains(active())).toBe(true);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"F2 goes to the newest toast",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await click(
				document.querySelector('header button[aria-label="More commands"]') as HTMLElement,
			);
			const copy = [...document.querySelectorAll<HTMLElement>('[role="menuitem"]')].find(
				(i) => i.textContent?.startsWith("Copy FetchXML"),
			);
			await click(copy);
			const toast = () => document.querySelector<HTMLElement>(".fui-Toast");
			await app.waitFor(() => !!toast());
			await act(async () => treeItem("name")?.focus());
			await press(active(), { key: "F2" });
			// Focus is on the toast's list item, which holds the toast
			await app.waitFor(() => active()?.getAttribute("role") === "listitem");
			expect(active()?.contains(toast())).toBe(true);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});
