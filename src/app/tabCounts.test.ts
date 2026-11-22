/**
 * The right pane's tab counts in the app (#233): the Results tab names the rows a run loaded, and
 * the Advisories tab the query's findings, coloured by the worst one.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, removeHost } from "../shared/testing/ui";

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
const tabNames = () =>
	[
		...document.querySelectorAll<HTMLElement>(
			'[role="tablist"][aria-label="Query views"] [role="tab"]',
		),
	].map((t) => t.getAttribute("aria-label") ?? t.textContent?.trim());

describe("the right pane's tab counts", () => {
	it(
		"Results names the rows a run loaded; Advisories names the query's findings",
		async () => {
			// A leading wildcard on like is performance advice; no unique order on a paged query is too
			const app = await renderApp({
				launch: {
					fetchXml:
						"<fetch count='50'><entity name='account'><attribute name='name'/><filter><condition attribute='name' operator='like' value='%a%'/></filter></entity></fetch>",
				},
			});
			await app.waitFor(() => !!button("Run", document));
			expect(tabNames()).toContain("Results");
			await app.waitFor(() => tabNames().some((n) => n?.startsWith("Advisories, ")));
			const advisories = tabNames().find((n) => n?.startsWith("Advisories, "));
			expect(advisories).toMatch(/^Advisories, \d+: /);

			await click(button("Run", document));
			await app.waitFor(() => /Rows: \d+/.test(app.status()));
			const rows = /Rows: (\d+)/.exec(app.status())?.[1];
			await app.waitFor(() => tabNames().includes(`Results, ${rows} rows`));
		},
		FLOW_TIMEOUT,
	);
});
