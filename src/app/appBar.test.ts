/**
 * The app bar in the app (#278): Run says why it can't run, as a focusable command with the reason
 * as its description: no entity yet, or a quick find view without a test value. In editor mode the
 * editor's query runs instead, so the quick find reason doesn't apply. Ctrl+Enter runs from the
 * tree, but not while a dialog is open.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { act } from "react";
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

const originalClipboard = Object.getOwnPropertyDescriptor(navigator, "clipboard");

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
	if (originalClipboard) Object.defineProperty(navigator, "clipboard", originalClipboard);
	else delete (navigator as { clipboard?: unknown }).clipboard;
});

/** Per test: the first render imports the whole app (about 5 s), and the suite runs in parallel */
const FLOW_TIMEOUT = 120000;

const accountQuery = `<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>`;
const quickFindQuery = `<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/><filter type='or'><condition attribute='name' operator='like' value='{0}'/></filter></entity></fetch>`;

const execute = () => button("Run", document);
const description = (element: Element | undefined) => {
	const id = element?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};

describe("the app bar's Run", () => {
	it(
		"with no entity yet: focusable, unavailable, 'Choose an entity first'",
		async () => {
			const app = await renderApp();
			await app.waitFor(() => !!execute());
			expect(execute()?.getAttribute("aria-disabled")).toBe("true");
			expect(execute()?.hasAttribute("disabled")).toBe(false);
			expect(description(execute())).toBe("Choose an entity first");
		},
		FLOW_TIMEOUT,
	);

	it(
		"a quick find view needs a test value, except in editor mode",
		async () => {
			const app = await renderApp({ launch: { fetchXml: quickFindQuery } });
			await app.waitFor(() => !!execute() && !!description(execute()));
			expect(execute()?.getAttribute("aria-disabled")).toBe("true");
			expect(description(execute())).toBe("Enter a quick find value first");

			// Editor mode with the same FetchXML in the editor: the editor's query runs after syncing
			Object.defineProperty(navigator, "clipboard", {
				configurable: true,
				value: { readText: async () => quickFindQuery },
			});
			await click(document.querySelector('input[role="switch"]'));
			await click(button("Paste", document));
			await app.waitFor(() => execute()?.getAttribute("aria-disabled") === null);
			expect(description(execute())).toBe("Run query (Ctrl+Enter)");
		},
		FLOW_TIMEOUT,
	);

	it(
		"Ctrl+Enter runs from the tree, but not while a dialog is open",
		async () => {
			const app = await renderApp({ launch: { fetchXml: accountQuery } });
			await app.waitFor(() => execute()?.getAttribute("aria-disabled") === null);
			const runs = () =>
				app.dataverse.calls.filter((c) => c.startsWith("queryData accounts?fetchXml="))
					.length;
			const ctrlEnter = async (target: Element | null) => {
				await act(async () => {
					target?.dispatchEvent(
						new KeyboardEvent("keydown", {
							key: "Enter",
							ctrlKey: true,
							bubbles: true,
						}),
					);
				});
				await flush(5);
			};

			// With the Settings drawer open, the keys belong to it
			await click(
				button("Settings", document) ?? document.querySelector('[aria-label="Settings"]'),
			);
			await app.waitFor(() => !!document.querySelector('[role="dialog"]'));
			await ctrlEnter(document.querySelector('[role="dialog"]'));
			expect(runs()).toBe(0);
			await act(async () => {
				document
					.querySelector('[role="dialog"]')
					?.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
			});
			await app.waitFor(() => !document.querySelector('[role="dialog"]'));

			await ctrlEnter(document.querySelector('[role="treeitem"]'));
			await app.waitFor(() => /Rows: 5\b/.test(app.status()));
			expect(runs()).toBe(1);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});
