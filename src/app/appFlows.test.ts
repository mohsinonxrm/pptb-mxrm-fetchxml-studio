/**
 * The whole tool, end to end against an in-memory Dataverse (#189, safety net for #301): what a
 * user sees when a query runs, pages, and changes entity. These tests read the screen, not the
 * code's structure, so they keep passing while AppShell is split into hooks.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, isDisabled, openDialog, removeHost, scopeField } from "../shared/testing/ui";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = (props: { value?: string; defaultValue?: string }) =>
		React.createElement("textarea", { "data-monaco": "", readOnly: true, value: props.value ?? props.defaultValue ?? "" });
	return { default: Stub, Editor: Stub, DiffEditor: Stub, loader: { config() {}, init: async () => ({}) }, useMonaco: () => null };
});

afterEach(() => {
	cleanup();
	removeHost();
});

/** Per test: the first render imports the whole app (about 5 s), and the suite runs in parallel */
const FLOW_TIMEOUT = 120000;

const accountQuery = (count = 50) => `<fetch count='${count}'><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>`;

describe("app: running a query", () => {
	it("a query from another tool loads into the tree, and Execute shows its rows", async () => {
		const app = await renderApp({ launch: { fetchXml: accountQuery() } });
		await app.waitFor(() => !!button("Execute", document));
		expect(treeRow("Fetch query")).toBeTruthy();
		expect(treeRow("name")).toBeTruthy();
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 5\b/.test(app.status()));
		expect(app.dataverse.calls.some((c) => c.startsWith("queryData accounts?fetchXml="))).toBe(true);
		app.expectClean();
	}, FLOW_TIMEOUT);

	it("a first page shows 'more available', and Retrieve all loads every page", async () => {
		const app = await renderApp({ launch: { fetchXml: accountQuery(2) } });
		await app.waitFor(() => !!button("Execute", document));
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 2 \(more available\)/.test(app.status()));

		await click(treeRow("Fetch query"));
		await app.waitFor(() => !!switchByLabel("Retrieve all records"));
		await click(switchByLabel("Retrieve all records")!);
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 5\b/.test(app.status()) && !/more available/.test(app.status()), 30000);
		const pages = app.dataverse.calls.filter((c) => c.startsWith("queryData accounts?fetchXml="));
		expect(pages.length).toBeGreaterThanOrEqual(1 + 3);
		app.expectClean();
	}, FLOW_TIMEOUT);

	it("Stop during Retrieve all keeps the rows already loaded and says so", async () => {
		let release!: () => void;
		const held = new Promise<void>((resolve) => (release = resolve));
		const app = await renderApp({
			launch: { fetchXml: accountQuery(2) },
			beforePage: (page) => (page === 2 ? held : undefined),
		});
		await app.waitFor(() => !!treeRow("Fetch query"));
		await click(treeRow("Fetch query"));
		await app.waitFor(() => !!switchByLabel("Retrieve all records"));
		await click(switchByLabel("Retrieve all records")!);
		await click(button("Execute", document));
		await app.waitFor(() => !!buttonOrNull("Stop"));
		await click(buttonOrNull("Stop")!);
		release();
		await app.waitFor(() => /Stopped after 2 rows/.test(document.body.textContent ?? ""));
		expect(app.status()).toMatch(/^Rows: 2\b/);
		const pages = app.dataverse.calls.filter((c) => c.startsWith("queryData accounts?fetchXml="));
		expect(pages.length).toBe(2);
		app.expectClean();
	}, FLOW_TIMEOUT);

	it("a query Dataverse rejects shows the error, and the tool stays usable", async () => {
		let fail = true;
		const app = await renderApp({
			launch: { fetchXml: accountQuery() },
			beforePage: () => {
				if (fail) throw new Error("0x80041103: 'account' entity doesn't contain attribute with Name = 'nme'.");
			},
		});
		await app.waitFor(() => !!button("Execute", document));
		await click(button("Execute", document));
		await app.waitFor(() => /doesn't contain attribute/.test(document.body.textContent ?? ""));
		expect(app.status()).toBe("");
		fail = false;
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 5\b/.test(app.status()));
		expect(document.body.textContent).not.toMatch(/doesn't contain attribute/);
		app.expectClean();
	}, FLOW_TIMEOUT);
});

describe("app: record actions", () => {
	it("deleting a selected row confirms first, names the target, deletes that record and re-runs", async () => {
		const app = await renderApp({ launch: { fetchXml: accountQuery() } });
		await app.waitFor(() => !!button("Execute", document));
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 5\b/.test(app.status()));

		// The first data row: "Account 01", ordered by name
		const firstRow = [...document.querySelectorAll<HTMLElement>('[role="grid"] [role="row"]')].find((r) => r.textContent?.includes("Account 01"));
		const checkbox = firstRow?.querySelector<HTMLInputElement>('input[type="checkbox"]');
		expect(checkbox).toBeTruthy();
		await click(checkbox!);
		await app.waitFor(() => !!buttonOrNull("Delete"));
		await click(buttonOrNull("Delete")!);

		await app.waitFor(() => !!openDialog() && !isDisabled(button("Delete")));
		expect(scopeField("Environment")).toContain("Test environment");
		expect(app.dataverse.deleted).toEqual([]);
		await click(button("Delete"));

		await app.waitFor(() => /Rows: 4\b/.test(app.status()));
		const accounts = app.dataverse.tables.find((t) => t.logicalName === "account")!;
		expect(app.dataverse.deleted).toHaveLength(1);
		expect(accounts.rows.map((r) => r.name)).toEqual(["Account 02", "Account 03", "Account 04", "Account 05"]);
		expect(document.body.textContent).not.toContain("Account 01");
		app.expectClean();
	}, FLOW_TIMEOUT);
});

function buttonOrNull(name: string): HTMLButtonElement | null {
	return [...document.querySelectorAll("button")].find((b) => b.textContent?.trim() === name) ?? null;
}

const VIEW_ID = "11111111-2222-3333-4444-555555555555";
const activeAccounts = {
	id: VIEW_ID,
	name: "Active accounts",
	entity: "account",
	type: "system" as const,
	fetchxml: "<fetch><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>",
	layoutxml: "<grid name='resultset' object='1' jump='name' select='1' icon='1' preview='1'><row name='result' id='accountid'><cell name='name' width='300'/></row></grid>",
};

describe("app: opening a view from another tool", () => {
	it("a viewRef loads the view, and Execute runs it by its id", async () => {
		const app = await renderApp({ views: [activeAccounts], launch: { viewRef: { id: VIEW_ID, entityLogicalName: "savedquery" } } });
		await app.waitFor(() => !!button("Execute", document) && (document.body.textContent ?? "").includes("Active accounts"));
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 5\b/.test(app.status()));
		expect(app.dataverse.calls).toContain(`queryData accounts?savedQuery=${VIEW_ID}`);
		expect(app.notifications).toEqual([]);
		app.expectClean();
	}, FLOW_TIMEOUT);

	it("a view with more pages follows Dataverse's @odata.nextLink to load the rest", async () => {
		const paged = { ...activeAccounts, fetchxml: activeAccounts.fetchxml.replace("<fetch>", "<fetch count='2'>") };
		const app = await renderApp({ views: [paged], launch: { viewRef: { id: VIEW_ID, entityLogicalName: "savedquery" } } });
		await app.waitFor(() => !!button("Execute", document) && (document.body.textContent ?? "").includes("Active accounts"));
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 2 \(more available\)/.test(app.status()));

		await click(treeRow("Fetch query"));
		await app.waitFor(() => !!switchByLabel("Retrieve all records"));
		await click(switchByLabel("Retrieve all records")!);
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 5\b/.test(app.status()) && !/more available/.test(app.status()), 30000);
		const viewPages = app.dataverse.calls.filter((c) => c.startsWith(`queryData accounts?savedQuery=${VIEW_ID}`));
		expect(viewPages.some((c) => c.endsWith("&$skiptoken=3"))).toBe(true);
		app.expectClean();
	}, FLOW_TIMEOUT);

	it("a viewRef to a view that doesn't exist tells the user, and leaves the tool usable", async () => {
		const missing = "99999999-2222-3333-4444-555555555555";
		const app = await renderApp({ views: [activeAccounts], launch: { viewRef: { id: missing, entityLogicalName: "savedquery" } } });
		await app.waitFor(() => app.notifications.length > 0);
		expect(app.notifications[0]).toMatchObject({ title: "Couldn't open the view", type: "error" });
		expect(app.notifications[0].body).toContain(missing);
		// Usable: nothing loaded, and a new query can start
		expect(app.dataverse.calls.some((c) => /\?(fetchXml|savedQuery)=/.test(c))).toBe(false);
		expect(isDisabled(button("New", document))).toBe(false);
		app.expectClean();
	}, FLOW_TIMEOUT);
});

/** The tree row whose own text is `text` (Fluent TreeItemLayout), for a real click on the label */
function treeRow(text: string): HTMLElement {
	const rows = [...document.querySelectorAll<HTMLElement>('[role="treeitem"]')];
	const row = rows.find((r) => r.querySelector(".fui-TreeItemLayout__main")?.textContent?.trim().startsWith(text));
	const target = row?.querySelector<HTMLElement>(".fui-TreeItemLayout__main") ?? row;
	if (!target) throw new Error(`No tree row "${text}". Rows: ${rows.map((r) => r.textContent?.slice(0, 30)).join(" | ")}`);
	return target;
}

/** A Fluent Switch's input, found by the text of its label */
function switchByLabel(text: string): HTMLInputElement | null {
	const label = [...document.querySelectorAll("label")].find((l) => l.textContent?.trim().startsWith(text));
	const id = label?.getAttribute("for");
	return id ? (document.getElementById(id) as HTMLInputElement | null) : null;
}
