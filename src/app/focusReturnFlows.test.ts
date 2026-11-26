/**
 * Focus comes back to what opened a dialog, drawer or popover once it closes (#279 PR 2, using
 * #186's focus return), with the real openers in the whole tool: Settings, Edit columns, the scope
 * chip, Open view, and the Delete record dialog, cancelled or done (#393: the query runs again
 * behind it, and Delete becomes Delete all).
 */
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, removeHost, stubOverflowWidths } from "../shared/testing/ui";

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
	"<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/><order attribute='name'/></entity></fetch>";

type App = Awaited<ReturnType<typeof renderApp>>;
const active = () => document.activeElement as HTMLElement | null;
const settle = (ms = 400) => act(() => new Promise<void>((resolve) => setTimeout(resolve, ms)));
const escape = () =>
	act(async () => {
		active()?.dispatchEvent(
			new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
		);
	});
/** A control by its accessible name or text, anywhere on the page */
const named = (name: string) =>
	[...document.querySelectorAll<HTMLElement>("button")].find(
		(b) => b.getAttribute("aria-label") === name || b.textContent?.trim() === name,
	);
const popup = () =>
	document.querySelector<HTMLElement>(
		'[role="dialog"]:not([aria-modal="false"]), [role="menu"], [role="alertdialog"]',
	);

/** The first-run scope tip traps focus: dismiss it as a user would */
const ready = async (app: App) => {
	await app.waitFor(() => !!document.querySelector('[role="treeitem"]'));
	const gotIt = named("Got it");
	if (gotIt) await click(gotIt);
	await app.waitFor(() => !document.querySelector('[role="dialog"][aria-modal="true"]'));
	await settle();
};
const runQuery = async (app: App) => {
	await click(button("Run", document));
	await app.waitFor(() => /Rows: \d/.test(app.status()));
	await settle();
};

/** Opens `opener` with a click, closes what it opened with Escape, and expects focus back on it */
const opensAndReturns = async (app: App, name: string) => {
	const opener = named(name);
	expect(opener, name).toBeTruthy();
	opener?.focus();
	await click(opener);
	await app.waitFor(() => !!popup());
	// Focus must be in what opened, or "back on the opener" proves nothing. A popover that traps focus
	// doesn't move it there in jsdom: move it to its first control, as Tab would
	if (!popup()!.contains(active())) {
		popup()!
			.querySelector<HTMLElement>("button, input, [tabindex]:not([tabindex='-1'])")
			?.focus();
	}
	expect(popup()!.contains(active()), `focus in what ${name} opened`).toBe(true);
	await escape();
	await app.waitFor(() => !popup());
	await settle(200);
	expect(active(), `focus back on ${name}`).toBe(opener);
};

describe("app: focus returns to the opener", () => {
	it(
		"Settings, the scope chip and Open view",
		async () => {
			// Room for every command: without layout, the bars would put them all in More
			stubOverflowWidths({ container: 4000 });
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await opensAndReturns(app, "Settings");
			const findChip = () =>
				[...document.querySelectorAll<HTMLElement>("button")].find((b) =>
					/^(Choose scope|Entity scope)/.test(b.getAttribute("aria-label") ?? ""),
				);
			await app.waitFor(() => !!findChip());
			const chip = [...document.querySelectorAll<HTMLElement>("button")].find((b) =>
				/^(Choose scope|Entity scope)/.test(b.getAttribute("aria-label") ?? ""),
			);
			await opensAndReturns(app, chip?.getAttribute("aria-label") ?? "Choose scope");
			await opensAndReturns(app, "Open view");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"Edit columns",
		async () => {
			// Room for every command: without layout, the bars would put them all in More
			stubOverflowWidths({ container: 4000 });
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await runQuery(app);
			await opensAndReturns(app, "Edit columns");
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);

	it(
		"the Delete record dialog: cancelled, and done (the query runs again, Delete becomes Delete all)",
		async () => {
			// Room for every command: without layout, the bars would put them all in More
			stubOverflowWidths({ container: 4000 });
			const app = await renderApp({ launch: { fetchXml: QUERY } });
			await ready(app);
			await runQuery(app);
			const rowBox = document.querySelectorAll<HTMLElement>(
				'[role="row"] input[type="checkbox"]',
			)[1];
			await click(rowBox);
			await settle(200);
			const del = named("Delete selected records");
			expect(del).toBeTruthy();
			// Cancelled
			del?.focus();
			await click(del);
			await app.waitFor(() => !!popup());
			await click(button("Cancel", document));
			await app.waitFor(() => !popup());
			await settle(200);
			expect(active()).toBe(del);
			// Done: it deletes, the query runs again, and Close returns focus to the same command
			await click(del);
			await app.waitFor(() => !!popup());
			await click(button("Delete"));
			await app.waitFor(() => !!button("Close"), 20000);
			await settle(300);
			await click(button("Close"));
			await app.waitFor(() => !popup());
			await settle(300);
			expect(active()).toBe(del);
			expect(del?.getAttribute("aria-label")).toBe("Delete all records from view");
		},
		FLOW_TIMEOUT,
	);
});
