/**
 * Open view in the app (#278 PR 4): the app bar opens a view of the query's entity, which loads like
 * a launched view (its FetchXML, layout and identity), after asking about unsaved changes. The left
 * pane no longer has a Load view row.
 */
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

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

const FLOW_TIMEOUT = 120000;
const accountQuery = `<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/></entity></fetch>`;
const activeAccounts = {
	id: "11111111-2222-3333-4444-555555555555",
	name: "Active accounts",
	entity: "account",
	type: "system" as const,
	fetchxml:
		"<fetch><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>",
	layoutxml:
		"<grid name='resultset' object='1' jump='name' select='1' icon='1' preview='1'><row name='result' id='accountid'><cell name='name' width='300'/><cell name='accountid' width='100'/></row></grid>",
};
const myAccounts = {
	...activeAccounts,
	id: "22222222-3333-4444-5555-666666666666",
	name: "My accounts",
	type: "personal" as const,
};

const appBar = () =>
	document.querySelector<HTMLElement>('[role="toolbar"][aria-label="Query commands"]');
const identity = () =>
	appBar()?.parentElement?.querySelector("strong, .fui-Body1Strong")?.textContent;
const options = () =>
	[...document.querySelectorAll<HTMLElement>('.fui-PopoverSurface [role="option"]')].map(
		(o) => o.querySelector("span > span")?.textContent,
	);
const option = (name: string) =>
	[...document.querySelectorAll<HTMLElement>('.fui-PopoverSurface [role="option"]')].find(
		(o) => o.querySelector("span > span")?.textContent === name,
	);
const alertButton = (label: string) =>
	[
		...(document
			.querySelector('[role="alertdialog"]')
			?.querySelectorAll<HTMLButtonElement>("button") ?? []),
	].find((b) => b.textContent?.trim() === label);

describe("Open view", () => {
	it(
		"opens a view of the query's entity: its FetchXML, layout and name",
		async () => {
			const app = await renderApp({
				views: [activeAccounts, myAccounts],
				launch: { fetchXml: accountQuery },
			});
			await app.waitFor(() => document.querySelectorAll('[role="treeitem"]').length > 0);
			// The left pane has no Load view row any more
			expect(document.body.textContent).not.toContain("Load view");

			await app.waitFor(() => {
				const open = button("Open view", appBar() ?? undefined);
				return !!open && open.getAttribute("aria-disabled") !== "true";
			});
			await click(button("Open view", appBar() ?? undefined));
			await app.waitFor(() => options().length === 2);
			expect(options()).toEqual(["My accounts", "Active accounts"]);

			await click(option("Active accounts"));
			// The launched query isn't a saved view: Open view asks first
			await app.waitFor(() => !!document.querySelector('[role="alertdialog"]'));
			await click(alertButton("Discard changes"));
			await app.waitFor(() => identity() === "Active accounts", 10000);
			await flush(5);
			// The view's FetchXML (it sorts by name; the launched query didn't)
			const xml = document.querySelector<HTMLTextAreaElement>("textarea[data-monaco]")?.value;
			expect(xml).toContain('<order attribute="name"');
			// Opened, unchanged: no badge
			expect(appBar()?.parentElement?.textContent).not.toContain("Unsaved changes");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});
