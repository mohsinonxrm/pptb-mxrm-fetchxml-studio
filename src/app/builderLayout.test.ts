/**
 * The builder pane's layout in the app (#234 PR 1): landmarks, the two keyboard splitters, and
 * Properties in the inspector drawer when the window is narrow (a selection bar names the node and
 * opens it; closing returns focus there).
 */
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, flush, removeHost } from "../shared/testing/ui";

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

const originalWidth = window.innerWidth;
afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
	Object.defineProperty(window, "innerWidth", { configurable: true, value: originalWidth });
});

const FLOW_TIMEOUT = 120000;
const query = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";
const builder = () => document.querySelector<HTMLElement>('aside[aria-label="Query builder"]');
const splitter = (label: string) =>
	document.querySelector<HTMLElement>(`[role="separator"][aria-label="${label}"]`);
const WIDTH = "Resize the builder pane (size is remembered)";
const TREE = "Resize the query tree (size is remembered)";
const key = async (element: HTMLElement | null, k: string) => {
	element?.focus();
	await act(async () => {
		element?.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true }));
	});
	await flush(2);
};
/** The tree row whose own text starts with `text` (Fluent TreeItemLayout), for a real click */
const treeItem = (text: string) => {
	const rows = [...document.querySelectorAll<HTMLElement>('[role="treeitem"]')];
	const row = rows.find((r) =>
		r.querySelector(".fui-TreeItemLayout__main")?.textContent?.trim().startsWith(text),
	);
	const target = row?.querySelector<HTMLElement>(".fui-TreeItemLayout__main") ?? row;
	if (!target)
		throw new Error(
			`No tree row "${text}". Rows: ${rows.map((r) => r.textContent?.slice(0, 30)).join(" | ")}`,
		);
	return target;
};

describe("the builder pane's layout", () => {
	it(
		"is a Query builder aside beside the Query output, resized with the keyboard",
		async () => {
			Object.defineProperty(window, "innerWidth", { configurable: true, value: 1280 });
			const app = await renderApp({ launch: { fetchXml: query } });
			await app.waitFor(() => !!builder());
			expect(document.querySelector('main[aria-label="Query output"]')).toBeTruthy();
			expect(builder()?.style.width).toBe("380px");

			const width = splitter(WIDTH);
			expect(width?.getAttribute("aria-controls")).toBe(builder()?.id);
			expect(width?.getAttribute("aria-valuenow")).toBe("380");
			await key(width, "ArrowRight");
			expect(builder()?.style.width).toBe("396px");

			const tree = splitter(TREE);
			const treeRegion = document.getElementById(tree?.getAttribute("aria-controls") ?? "");
			expect(treeRegion?.style.height).toBe("240px");
			await key(tree, "ArrowDown");
			expect(treeRegion?.style.height).toBe("256px");
			// No old mouse-only handles
			expect(document.querySelector('[title="Drag to resize"]')).toBeNull();
		},
		FLOW_TIMEOUT,
	);

	it(
		"puts Properties in a drawer when the window is narrow, opened from the selection bar",
		async () => {
			Object.defineProperty(window, "innerWidth", { configurable: true, value: 900 });
			const app = await renderApp({ launch: { fetchXml: query } });
			await app.waitFor(() => document.querySelectorAll('[role="treeitem"]').length > 0);
			// No docked Properties and no tree splitter; the selection bar instead
			expect(splitter(TREE)).toBeNull();
			const bar = () =>
				document.querySelector<HTMLElement>('[role="region"][aria-label="Selected node"]');
			expect(bar()).toBeTruthy();

			await click(treeItem("account"));
			await app.waitFor(() => !!bar()?.textContent?.includes("Entity · account"));
			const edit = button("Edit properties", bar() ?? undefined);
			expect(edit?.getAttribute("aria-disabled")).not.toBe("true");
			await click(edit);
			const drawer = () =>
				document.querySelector<HTMLElement>('[aria-label="Properties: Entity · account"]');
			await app.waitFor(() => !!drawer());
			// The drawer sits in the output, under the tabs
			expect(
				document.querySelector('main[aria-label="Query output"]')?.contains(drawer()),
			).toBe(true);
			expect(drawer()?.textContent).toContain("account");

			await click(document.querySelector('button[aria-label="Close properties"]'));
			await app.waitFor(() => !drawer());
			expect(document.activeElement).toBe(button("Edit properties", bar() ?? undefined));

			// Escape from inside the drawer closes it the same way
			await click(button("Edit properties", bar() ?? undefined));
			await app.waitFor(() => !!drawer());
			const inside = drawer()?.querySelector<HTMLElement>("button, input") ?? drawer();
			await key(inside ?? null, "Escape");
			await app.waitFor(() => !drawer());
			expect(document.activeElement).toBe(button("Edit properties", bar() ?? undefined));
		},
		FLOW_TIMEOUT,
	);
});
