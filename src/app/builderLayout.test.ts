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
const linkedQuery =
	"<fetch><entity name='account'><attribute name='name'/><link-entity name='systemuser' from='systemuserid' to='createdby' link-type='outer' alias='cb'><attribute name='fullname'/></link-entity></entity></fetch>";
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

			// Widening the window while focus is in the drawer docks Properties; focus goes to the
			// selected tree item instead of the page
			await click(button("Edit properties", bar() ?? undefined));
			await app.waitFor(() => !!drawer());
			drawer()?.querySelector<HTMLElement>("button")?.focus();
			Object.defineProperty(window, "innerWidth", { configurable: true, value: 1280 });
			await act(async () => {
				window.dispatchEvent(new Event("resize"));
			});
			await app.waitFor(() => !!splitter(TREE));
			expect(drawer()).toBeNull();
			expect(document.activeElement?.closest('[role="treeitem"]')?.textContent).toContain(
				"account",
			);

			// New clears the selection: the drawer closes instead of showing nothing
			Object.defineProperty(window, "innerWidth", { configurable: true, value: 900 });
			await act(async () => {
				window.dispatchEvent(new Event("resize"));
			});
			await app.waitFor(() => !!bar());
			await click(button("Edit properties", bar() ?? undefined));
			await app.waitFor(() => !!drawer());
			const appBar = document.querySelector<HTMLElement>(
				'[role="toolbar"][aria-label="Query commands"]',
			);
			await click(button("New", appBar ?? undefined));
			const discard = () =>
				[
					...(document
						.querySelector('[role="alertdialog"]')
						?.querySelectorAll<HTMLButtonElement>("button") ?? []),
				].find((b) => b.textContent?.trim() === "Discard changes");
			await app.waitFor(() => !!discard() || !document.querySelector('[role="treeitem"]'));
			if (discard()) await click(discard());
			await app.waitFor(() => !document.querySelector('[role="treeitem"]'));
			await app.waitFor(() => !drawer());
			expect(bar()?.textContent).toContain("Nothing selected");
			expect(document.querySelector('[aria-label^="Properties:"]')).toBeNull();
		},
		FLOW_TIMEOUT,
	);

	it(
		"collapses to a rail that peeks at the tree or Properties, and Pin docks it again",
		async () => {
			Object.defineProperty(window, "innerWidth", { configurable: true, value: 1280 });
			const app = await renderApp({ launch: { fetchXml: linkedQuery } });
			await app.waitFor(() => document.querySelectorAll('[role="treeitem"]').length > 0);
			const named = (name: string) =>
				document.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`);
			const rail = () =>
				document.querySelector<HTMLElement>('nav[aria-label="Builder pane"]');
			const peek = (label: string) =>
				document.querySelector<HTMLElement>(`aside[aria-label="${label}"]`);

			// A branch opened while docked stays open through collapsing and peeking (one tree)
			const cbRow = () =>
				[...document.querySelectorAll<HTMLElement>('[role="treeitem"]')].find((r) =>
					r
						.querySelector(".fui-TreeItemLayout__main")
						?.textContent?.trim()
						.startsWith("cb"),
				);
			if (cbRow()?.getAttribute("aria-expanded") !== "true") await click(treeItem("cb"));
			await app.waitFor(() => cbRow()?.getAttribute("aria-expanded") === "true");

			// Collapse sits beside the scope chip; it leaves the rail, with focus on Expand
			expect(rail()).toBeNull();
			await click(named("Collapse the builder pane"));
			await app.waitFor(() => !!rail());
			expect(builder()?.style.display).toBe("none");
			expect(splitter(WIDTH)).toBeNull();
			expect(document.activeElement).toBe(named("Expand the builder pane"));

			// Query tree opens the tree in a peek, focus inside it on the tree
			await click(named("Query tree"));
			await app.waitFor(() => !!peek("Query tree"));
			expect(named("Query tree")?.getAttribute("aria-pressed")).toBe("true");
			expect(named("Query tree")?.getAttribute("aria-controls")).toBe(peek("Query tree")?.id);
			expect(peek("Query tree")?.contains(document.activeElement)).toBe(true);
			expect(document.activeElement?.getAttribute("role")).toBe("treeitem");
			expect(cbRow()?.getAttribute("aria-expanded")).toBe("true");
			// The scope isn't checked again: no access spinner
			expect(peek("Query tree")?.textContent).not.toContain("Checking access");
			// Escape closes it and returns focus to its rail button
			await key(document.activeElement as HTMLElement, "Escape");
			await app.waitFor(() => !peek("Query tree"));
			expect(document.activeElement).toBe(named("Query tree"));

			// Properties peek: select a node first, then open its properties; focus on the first
			// property, not Pin
			await click(named("Query tree"));
			await app.waitFor(() => !!peek("Query tree"));
			await click(treeItem("cb"));
			await click(named("Properties"));
			await app.waitFor(() => !!peek("Properties"));
			expect(peek("Query tree")).toBeNull();
			expect(peek("Properties")?.textContent).toContain("Link-entity");
			// The tree stays mounted, hidden
			expect(document.getElementById("fxs-query-tree")?.style.display).toBe("none");
			expect(peek("Properties")?.contains(document.activeElement)).toBe(true);
			expect(document.activeElement).not.toBe(named("Pin the builder pane open"));
			expect(document.activeElement?.closest("[data-pane-command]")).toBeNull();
			// Its rail button again closes it
			await click(named("Properties"));
			await app.waitFor(() => !peek("Properties"));

			// A click on the output closes a peek
			await click(named("Query tree"));
			await app.waitFor(() => !!peek("Query tree"));
			await act(async () => {
				document
					.querySelector('main[aria-label="Query output"]')
					?.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true }));
			});
			await app.waitFor(() => !peek("Query tree"));

			// Pin docks the pane again, with focus on Collapse
			await click(named("Properties"));
			await app.waitFor(() => !!peek("Properties"));
			await click(named("Pin the builder pane open"));
			await app.waitFor(() => !!builder() && builder()?.style.display !== "none");
			expect(rail()).toBeNull();
			expect(splitter(WIDTH)).toBeTruthy();
			expect(document.activeElement).toBe(named("Collapse the builder pane"));

			// Expand does the same from the rail
			await click(named("Collapse the builder pane"));
			await app.waitFor(() => !!rail());
			await click(named("Expand the builder pane"));
			await app.waitFor(() => !!builder() && builder()?.style.display !== "none");
			expect(document.activeElement).toBe(named("Collapse the builder pane"));
		},
		FLOW_TIMEOUT,
	);
});
