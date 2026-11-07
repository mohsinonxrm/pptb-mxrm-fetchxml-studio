/**
 * The whole tool, end to end against an in-memory Dataverse (#189, safety net for #301): what a
 * user sees when a query runs, pages, and changes entity. These tests read the screen, not the
 * code's structure, so they keep passing while AppShell is split into hooks.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, removeHost } from "../shared/testing/ui";

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

const accountQuery = (count = 50) => `<fetch count='${count}'><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>`;

describe("app: running a query", () => {
	it("a query from another tool loads into the tree, and Execute shows its rows", async () => {
		const app = await renderApp({ launch: { fetchXml: accountQuery() } });
		await app.waitFor(() => !!button("Execute", document));
		await click(button("Execute", document));
		await app.waitFor(() => /Rows: 5/.test(app.status()));
		expect(app.dataverse.calls.some((c) => c.startsWith("queryData accounts?fetchXml="))).toBe(true);
		expect(app.dataverse.unknown).toEqual([]);
	}, 60000);
});
