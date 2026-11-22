/**
 * New in the app (#278 PR 2): it asks before it discards anything not saved, including FetchXML in
 * the editor (with or without a tree, and text typed a moment ago), and "Save and start new" only
 * starts new after its own save, never after a later one.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import {
	button,
	cleanup,
	click,
	dialogInput,
	flush,
	removeHost,
	typeInto,
} from "../shared/testing/ui";

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

const FLOW_TIMEOUT = 120000;
const accountQuery = `<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>`;
const otherQuery = `<fetch><entity name='account'><attribute name='name'/><attribute name='statecode'/></entity></fetch>`;
const VIEW_ID = "11111111-2222-3333-4444-555555555555";
const activeAccounts = {
	id: VIEW_ID,
	name: "Active accounts",
	entity: "account",
	type: "system" as const,
	fetchxml:
		"<fetch><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>",
	layoutxml:
		"<grid name='resultset' object='1' jump='name' select='1' icon='1' preview='1'><row name='result' id='accountid'><cell name='name' width='300'/><cell name='accountid' width='100'/></row></grid>",
};

const appBar = () =>
	document.querySelector<HTMLElement>('[role="toolbar"][aria-label="Query commands"]');
const treeItems = () => document.querySelectorAll('[role="treeitem"]').length;
const alert = () => document.querySelector('[role="alertdialog"]');
const saveReady = () => {
	const save = button(/^Save/, appBar() ?? undefined);
	return !!save && save.getAttribute("aria-disabled") !== "true";
};
const editorText = () =>
	document.querySelector<HTMLTextAreaElement>("textarea[data-monaco]")?.value;
async function pasteIntoEditor(xml: string) {
	Object.defineProperty(navigator, "clipboard", {
		configurable: true,
		value: { readText: async () => xml },
	});
	const toggle = document.querySelector<HTMLInputElement>('input[role="switch"]');
	if (!toggle?.checked) await click(toggle);
	await click(button("Paste", document));
}
const alertButton = (label: string) =>
	[...(alert()?.querySelectorAll<HTMLButtonElement>("button") ?? [])].find(
		(b) => b.textContent?.trim() === label,
	);

describe("New with FetchXML in the editor", () => {
	it(
		"asks when the editor has FetchXML and there's no tree",
		async () => {
			const app = await renderApp();
			await app.waitFor(() => !!button("Run", document));
			await pasteIntoEditor(accountQuery);
			await click(button("New", appBar() ?? undefined));
			await app.waitFor(() => !!alert());
			await click(alertButton("Cancel"));
			expect(editorText()).toBe(accountQuery);
		},
		FLOW_TIMEOUT,
	);

	it(
		"asks for text typed a moment ago, before the badge has caught up",
		async () => {
			const app = await renderApp({
				views: [activeAccounts],
				launch: { viewRef: { id: VIEW_ID, entityLogicalName: "savedquery" } },
			});
			await app.waitFor(() => treeItems() > 0 && saveReady());
			await pasteIntoEditor(otherQuery);
			// No wait: the badge updates 300 ms after typing stops, New mustn't depend on it
			await click(button("New", appBar() ?? undefined));
			await flush(3);
			expect(alert()).toBeTruthy();
			expect(treeItems()).toBeGreaterThan(0);
		},
		FLOW_TIMEOUT,
	);
});

describe("Save and start new", () => {
	it(
		"a save that didn't open (invalid FetchXML) doesn't make a later save start new",
		async () => {
			const app = await renderApp({ launch: { fetchXml: accountQuery } });
			await app.waitFor(() => treeItems() > 0 && saveReady());
			// Not well-formed: the editor sync stops Save before its dialog opens
			await pasteIntoEditor(
				"<fetch><entity name='account'><attribute name='name'/></entity>",
			);
			await click(button("New", appBar() ?? undefined));
			await app.waitFor(() => !!alert());
			await click(alertButton("Save and start new"));
			await flush(10);
			expect(document.querySelector('[role="dialog"]')).toBeNull();

			// Later, an ordinary Save as out of editor mode
			await click(document.querySelector('input[role="switch"]'));
			await click(button("Save", appBar() ?? undefined));
			await app.waitFor(() => !!document.querySelector('[role="dialog"]'));
			await typeInto(dialogInput(), "My view");
			await click(
				[...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(
					(b) => b.textContent?.trim() === "Save",
				),
			);
			await app.waitFor(() => !document.querySelector('[role="dialog"]'), 10000);
			await flush(10);
			expect(treeItems()).toBeGreaterThan(0);
		},
		FLOW_TIMEOUT,
	);

	it(
		"saves, then starts new",
		async () => {
			const app = await renderApp({ launch: { fetchXml: accountQuery } });
			await app.waitFor(() => treeItems() > 0 && saveReady());
			await click(button("New", appBar() ?? undefined));
			await app.waitFor(() => !!alert());
			await click(alertButton("Save and start new"));
			await app.waitFor(() => !!document.querySelector('[role="dialog"]'));
			expect(treeItems()).toBeGreaterThan(0);
			await typeInto(dialogInput(), "Saved before new");
			await click(
				[...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(
					(b) => b.textContent?.trim() === "Save",
				),
			);
			await app.waitFor(() => treeItems() === 0, 10000);
		},
		FLOW_TIMEOUT,
	);
});

describe("Save on a personal view you own (#278 PR 3)", () => {
	const PERSONAL_ID = "22222222-3333-4444-5555-666666666666";
	const myView = {
		...activeAccounts,
		id: PERSONAL_ID,
		name: "My accounts",
		type: "personal" as const,
	};
	const saveHint = () => {
		const save = button("Save", appBar() ?? undefined);
		const id = save?.getAttribute("aria-describedby");
		return id ? document.getElementById(id)?.textContent?.trim() : undefined;
	};

	it(
		"saves the changes to it, also after the editor's FetchXML changed the query",
		async () => {
			const app = await renderApp({
				views: [myView],
				launch: { viewRef: { id: PERSONAL_ID, entityLogicalName: "userquery" } },
			});
			await app.waitFor(() => treeItems() > 0 && saveReady());
			await app.waitFor(() => saveHint() === 'Save changes to "My accounts"');

			// Change the query in the editor: the view is still the one Save overwrites
			await pasteIntoEditor(otherQuery);
			await click(button("Save", appBar() ?? undefined));
			await app.waitFor(() => !!document.querySelector('[role="dialog"]'), 10000);
			const text = document.querySelector('[role="dialog"]')?.textContent ?? "";
			expect(text).toContain('The personal view "My accounts" (overwritten)');
			expect(text).not.toContain("Save as new view");
		},
		FLOW_TIMEOUT,
	);

	it(
		"after Save as a new personal view, Save saves changes to it: it's yours",
		async () => {
			const app = await renderApp({ launch: { fetchXml: accountQuery } });
			await app.waitFor(() => treeItems() > 0 && saveReady());
			await click(button("Save", appBar() ?? undefined));
			await app.waitFor(() => !!document.querySelector('[role="dialog"]'));
			await typeInto(dialogInput(), "Fresh view");
			await click(
				[...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(
					(b) => b.textContent?.trim() === "Save",
				),
			);
			await app.waitFor(() => !document.querySelector('[role="dialog"]'), 10000);
			// Its own view: no access check, no warning
			await app.waitFor(() => saveHint() === 'Save changes to "Fresh view"', 10000);
			expect(app.dataverse.calls.some((c) => c.includes("RetrievePrincipalAccess"))).toBe(
				false,
			);
		},
		FLOW_TIMEOUT,
	);
});
