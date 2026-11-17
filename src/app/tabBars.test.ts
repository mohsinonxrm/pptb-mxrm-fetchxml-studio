/**
 * The bars in the tabs (#277 PR 3): a run that failed is a bar in the Results tab with Try again and
 * Copy details; problems the editor-mode checks of Execute find are the FetchXML editor's own bars
 * (one per condition, owner decision D3 (a)), and Execute switches to the FetchXML tab to show them.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, removeHost, toasts } from "../shared/testing/ui";

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

/** jsdom's own navigator.clipboard (none), put back after a test replaces it */
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
/** Not well-formed */
const brokenXml = "<fetch><entity name='account'>";
/** Parses, but Dataverse might reject it: an operator that isn't documented */
const mightRejectXml = `<fetch count='50'><entity name='account'><attribute name='name'/><attribute name='accountid'/><filter><condition attribute='name' operator='bogus' value='x'/></filter></entity></fetch>`;

const bars = () =>
	[...document.querySelectorAll(".fui-MessageBar")].map((b) =>
		(b.textContent ?? "").replace(/\s+/g, " ").trim(),
	);
const barsTitled = (title: string) => bars().filter((b) => b.startsWith(title));
const TAB_NAMES = ["FetchXML", "LayoutXML", "Results", "Advisories", "Code"];
/** A tab's name (Fluent can render the label twice, to reserve the selected label's width) */
const tabName = (t: Element | null | undefined) =>
	TAB_NAMES.find((n) => t?.textContent?.trim().startsWith(n));
const selectedTab = () => tabName(document.querySelector('[role="tab"][aria-selected="true"]'));
const tab = (name: string) =>
	[...document.querySelectorAll('[role="tab"]')].find((t) => tabName(t) === name);

/** Turns editor mode on and pastes `xml` into the editor (the Monaco stub can't be typed into) */
async function editorWith(xml: string) {
	Object.defineProperty(navigator, "clipboard", {
		configurable: true,
		value: { readText: async () => xml },
	});
	await click(tab("FetchXML"));
	await click(document.querySelector('input[role="switch"]'));
	await click(button("Paste", document));
}

describe("the Results tab", () => {
	it(
		"a run that failed: 'The query didn't run' there only, Copy details, and Try again runs it",
		async () => {
			let fail = true;
			const app = await renderApp({
				launch: { fetchXml: accountQuery },
				beforePage: () => {
					if (fail)
						throw new Error(
							"0x80041103: 'account' entity doesn't contain attribute with Name = 'nme'.",
						);
				},
			});
			await app.waitFor(() => !!button("Execute", document));
			await click(button("Execute", document));
			await app.waitFor(() => barsTitled("The query didn't run").length > 0);
			expect(selectedTab()).toBe("Results");
			expect(barsTitled("The query didn't run")).toEqual([
				expect.stringContaining("doesn't contain attribute"),
			]);

			// Not above the other tabs
			await click(tab("FetchXML"));
			expect(barsTitled("The query didn't run")).toEqual([]);
			await click(tab("Results"));

			await click(button("Copy details", document));
			expect(toasts()).toEqual(["Copied error details."]);

			fail = false;
			await click(button("Try again", document));
			await app.waitFor(() => /Rows: 5\b/.test(app.status()));
			expect(barsTitled("The query didn't run")).toEqual([]);
			app.expectClean();
		},
		FLOW_TIMEOUT,
	);
});

describe("the FetchXML tab: the editor owns its problems", () => {
	it(
		"Execute with broken XML shows one error bar in the FetchXML tab, and Validate doesn't add a second",
		async () => {
			const app = await renderApp({ launch: { fetchXml: accountQuery } });
			await app.waitFor(() => !!button("Execute", document));
			await editorWith(brokenXml);
			await click(tab("Results"));

			await click(button("Execute", document));
			expect(selectedTab()).toBe("FetchXML");
			expect(barsTitled("Syntax error")).toHaveLength(1);
			expect(bars()).toHaveLength(1);

			await click(button("Validate", document));
			expect(bars()).toHaveLength(1);
			expect(barsTitled("Syntax error")).toHaveLength(1);

			// Opening the tab again doesn't clear it: that's where it gets fixed
			await click(tab("Results"));
			await click(tab("FetchXML"));
			expect(barsTitled("Syntax error")).toHaveLength(1);

			// Nothing ran
			expect(app.dataverse.calls.some((c) => c.startsWith("queryData accounts"))).toBe(false);
		},
		FLOW_TIMEOUT,
	);

	it(
		"might reject: one warning with Execute anyway, which runs the query and removes it",
		async () => {
			const app = await renderApp({ launch: { fetchXml: accountQuery } });
			await app.waitFor(() => !!button("Execute", document));
			await editorWith(mightRejectXml);
			await click(button("Validate", document));
			expect(barsTitled("Dataverse might reject this query")).toHaveLength(1);

			await click(tab("Results"));
			await click(button("Execute", document));
			expect(selectedTab()).toBe("FetchXML");
			expect(barsTitled("Dataverse might reject this query")).toEqual([
				expect.stringContaining("Execute anyway"),
			]);

			await click(button("Execute anyway", document));
			// Loading it into the tree may need the changes confirmed first
			const confirm = [...document.querySelectorAll('[role="dialog"] button')].find((b) =>
				/^(Execute|Continue|Load)/.test(b.textContent?.trim() ?? ""),
			);
			if (confirm) await click(confirm);
			await app.waitFor(() =>
				app.dataverse.calls.some((c) => c.startsWith("queryData accounts")),
			);
			expect(barsTitled("Dataverse might reject this query")).toEqual([]);
		},
		FLOW_TIMEOUT,
	);

	it(
		"turning editor mode off clears the editor's bars",
		async () => {
			const app = await renderApp({ launch: { fetchXml: accountQuery } });
			await app.waitFor(() => !!button("Execute", document));
			await editorWith(brokenXml);
			await click(button("Validate", document));
			expect(bars()).toHaveLength(1);
			await click(document.querySelector('input[role="switch"]'));
			expect(bars()).toEqual([]);
		},
		FLOW_TIMEOUT,
	);
});
