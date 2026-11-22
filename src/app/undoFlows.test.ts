/**
 * Undo and Redo in the whole tool (#192): the tree toolbar names the step, Ctrl+Z on a tree row and
 * the Undo button undo it, Ctrl+Y redoes it, and a text field keeps Ctrl+Z for itself. A query from
 * another tool is where the history starts, so there's nothing to undo after it loads.
 */
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { cleanup, click, removeHost } from "../shared/testing/ui";

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

/** A tree item by the start of its label */
const treeItem = (text: string) =>
	[...document.querySelectorAll<HTMLElement>('[role="treeitem"]')].find((r) =>
		r.querySelector(".fui-TreeItemLayout__main")?.textContent?.trim().startsWith(text),
	);
const treeToolbar = () =>
	document.querySelector<HTMLElement>('[role="toolbar"][aria-label="Query tree commands"]');
const historyButton = (verb: "Undo" | "Redo") =>
	[...(treeToolbar()?.querySelectorAll<HTMLButtonElement>("button") ?? [])].find((b) =>
		b.getAttribute("aria-label")?.startsWith(`${verb}`),
	);
/** The row's Delete button: Fluent shows a row's actions while it has focus */
const deleteRow = async (text: string) => {
	await act(async () => treeItem(text)?.focus());
	await click(treeItem(text)?.querySelector('button[aria-label="Delete"]'));
};
/** What Fluent's announcer last put in its live region (jsdom has no innerText, so read the property it sets) */
const liveText = () =>
	[...document.querySelectorAll<HTMLElement>("[aria-live] span")]
		.map((span) => span.innerText ?? span.textContent ?? "")
		.join(" ");
const nameOf = (verb: "Undo" | "Redo") => historyButton(verb)?.getAttribute("aria-label");
const press = (target: Element | null | undefined, init: KeyboardEventInit) =>
	act(async () => {
		target?.dispatchEvent(
			new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init }),
		);
	});

describe("app: undo and redo", () => {
	it(
		"the toolbar names the step; Ctrl+Z, Ctrl+Y and the Undo button undo and redo it",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await app.waitFor(() => !!treeItem("name") && !!treeToolbar());
			expect(nameOf("Undo")).toBe("Undo: Nothing to undo");
			expect(nameOf("Redo")).toBe("Redo: Nothing to redo");

			await deleteRow("name");
			await app.waitFor(() => !treeItem("name"));
			expect(nameOf("Undo")).toBe("Undo: Delete attribute (Ctrl+Z)");

			// Ctrl+Z on a tree row brings it back, selected
			await press(treeItem("Fetch query"), { key: "z", ctrlKey: true });
			await app.waitFor(() => !!treeItem("name"));
			expect(treeItem("name")?.getAttribute("aria-selected")).toBe("true");
			expect(nameOf("Redo")).toBe("Redo: Delete attribute (Ctrl+Y)");
			expect(nameOf("Undo")).toBe("Undo: Nothing to undo");

			// Ctrl+Y deletes it again; the button undoes that
			await press(treeItem("Fetch query"), { key: "y", ctrlKey: true });
			await app.waitFor(() => !treeItem("name"));
			await click(historyButton("Undo"));
			await app.waitFor(() => !!treeItem("name"));

			// Ctrl+Y on the focused row removes it: focus goes to the selected row, not the page
			await act(async () => treeItem("name")?.focus());
			await press(treeItem("name"), { key: "y", ctrlKey: true });
			await app.waitFor(() => !treeItem("name"));
			await app.waitFor(() => document.activeElement?.getAttribute("role") === "treeitem");
			expect(document.activeElement?.getAttribute("aria-selected")).toBe("true");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"screen readers hear what was undone, through the app's live region",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await app.waitFor(() => !!treeItem("name") && !!treeToolbar());
			await deleteRow("name");
			await app.waitFor(() => !treeItem("name"));
			await click(historyButton("Undo"));
			// Fluent's announcer writes the message into a live region for a moment
			await app.waitFor(() => liveText().includes("Undone: Delete attribute"), 5000);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"Ctrl+Z in a text field is the field's own",
		async () => {
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await app.waitFor(() => !!treeItem("name") && !!treeToolbar());
			await deleteRow("name");
			await app.waitFor(() => !treeItem("name"));
			const field = document.querySelector<HTMLInputElement>(
				'input[type="text"], input:not([type])',
			);
			expect(field).toBeTruthy();
			await press(field, { key: "z", ctrlKey: true });
			expect(treeItem("name")).toBeUndefined();
			expect(nameOf("Undo")).toBe("Undo: Delete attribute (Ctrl+Z)");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});
