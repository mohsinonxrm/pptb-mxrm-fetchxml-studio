/**
 * The Code tab (#233): languages in the owner's order, C# QueryExpression first (C# FetchExpression
 * while not connected), the language kept by the parent, where the code comes from, SQL's preview
 * note, and a conversion that failed with Try again and Copy details. The connection and the
 * Dataverse conversions are mocked; Monaco is a textarea.
 */
import { act, createElement, useState } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	button,
	cleanup,
	click,
	flush,
	render,
	stubOverflowWidths,
	toasts,
} from "../../../../shared/testing/ui";
import { CodePanel, type CodeLanguage } from "./CodePanel";
import { CODE_LANGUAGE_LABELS, CODE_LANGUAGES } from "./codePanelRules";
import { createConversionCache } from "./serverConversions";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = (props: { value?: string }) =>
		React.createElement("textarea", {
			"data-monaco": "",
			readOnly: true,
			value: props.value ?? "",
		});
	return { default: Stub, Editor: Stub, loader: { config() {}, init: async () => ({}) } };
});

const host = vi.hoisted(() => ({ connected: true }));
vi.mock("../../../../shared/hooks/usePptbContext", () => ({
	usePptbContext: () => ({
		theme: "light",
		connected: host.connected,
		environmentUrl: host.connected ? "https://test.crm.dynamics.com/" : undefined,
	}),
}));

const api = vi.hoisted(() => ({ fetchXmlToQueryExpression: vi.fn(), fetchXmlToSQL: vi.fn() }));
vi.mock("../../api/pptbClient", async (importOriginal) => ({
	...(await importOriginal<typeof import("../../api/pptbClient")>()),
	...api,
}));
vi.mock("../../api/dataverseMetadata", async (importOriginal) => ({
	...(await importOriginal<typeof import("../../api/dataverseMetadata")>()),
	loadEntityMetadata: async () => ({ EntitySetName: "accounts" }),
}));

const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";
const QE = { EntityName: "account", ColumnSet: { Columns: ["name"], AllColumns: false } };

/** The parent keeps the language, as PreviewTabs does */
function Code({ initial = null }: { initial?: CodeLanguage | null }) {
	const [language, setLanguage] = useState<CodeLanguage | null>(initial);
	return createElement(CodePanel, { fetchXml: XML, language, onLanguageChange: setLanguage });
}

const tabList = () =>
	document.querySelector<HTMLElement>('[role="tablist"][aria-label="Code languages"]');
const tabs = () => [...(tabList()?.querySelectorAll<HTMLElement>('[role="tab"]') ?? [])];
/** A tab's label: its visible content (Fluent keeps a hidden copy to reserve the bold width) */
const labelOf = (t: HTMLElement | undefined) =>
	t?.querySelector(".fui-Tab__content")?.textContent?.trim();
const selectedTab = () => tabs().find((t) => t.getAttribute("aria-selected") === "true");
const tab = (label: string) => tabs().find((t) => labelOf(t) === label);
const code = () => document.querySelector<HTMLTextAreaElement>("textarea[data-monaco]")?.value;
const text = () => (document.body.textContent ?? "").replace(/\s+/g, " ");
/** Waits out the conversion delay (500 ms after typing stops) and the request */
const settle = async () => {
	await new Promise((resolve) => setTimeout(resolve, 600));
	await flush(10);
};

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
	stubOverflowWidths({ container: 3000 });
	host.connected = true;
	api.fetchXmlToQueryExpression.mockResolvedValue(QE);
	api.fetchXmlToSQL.mockResolvedValue("SELECT name FROM account");
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
	api.fetchXmlToQueryExpression.mockReset();
	api.fetchXmlToSQL.mockReset();
});

describe("the Code tab's languages", () => {
	it("are in the owner's order, text only", async () => {
		await render(createElement(Code));
		expect(tabs().map(labelOf)).toEqual([
			"C# QueryExpression",
			"C# FetchExpression",
			"Web API",
			"Xrm.WebApi",
			"Power Automate",
			"pac CLI",
			"SQL (TDS)",
		]);
		expect(CODE_LANGUAGES.map((l) => CODE_LANGUAGE_LABELS[l])[0]).toBe("C# QueryExpression");
		// No badges or Tooltip icons inside the tabs
		for (const t of tabs()) expect(t.querySelector(".fui-Badge, svg")).toBeNull();
	});

	it("open on C# QueryExpression, from FetchXmlToQueryExpression, which the note names", async () => {
		await render(createElement(Code));
		expect(labelOf(selectedTab())).toBe("C# QueryExpression");
		await settle();
		expect(api.fetchXmlToQueryExpression).toHaveBeenCalledTimes(1);
		expect(code()).toContain('new QueryExpression("account")');
		expect(text()).toContain("From FetchXmlToQueryExpression");
	});

	it("open on C# FetchExpression while not connected, and say why", async () => {
		host.connected = false;
		await render(createElement(Code));
		expect(labelOf(selectedTab())).toBe("C# FetchExpression");
		expect(text()).toContain("Connect to an environment to generate QueryExpression.");
		expect(code()).toContain("FetchExpression");
		expect(api.fetchXmlToQueryExpression).not.toHaveBeenCalled();
		// Picked anyway: it says it needs a connection
		await click(tab("C# QueryExpression"));
		expect(text()).toContain(
			"Connect to a Dataverse environment to generate C# QueryExpression.",
		);
		expect(text()).not.toContain("Connect to an environment to generate QueryExpression.");
	});

	it("keep the language the parent keeps", async () => {
		await render(createElement(Code, { initial: "pac" }));
		expect(labelOf(selectedTab())).toBe("pac CLI");
		expect(api.fetchXmlToQueryExpression).not.toHaveBeenCalled();
		await click(tab("Web API"));
		expect(labelOf(selectedTab())).toBe("Web API");
		expect(text()).not.toContain("From Fetch");
	});
});

describe("SQL", () => {
	it("says it's a preview, with Learn more, and where it comes from", async () => {
		await render(createElement(Code, { initial: "sql" }));
		await settle();
		expect(code()).toBe("SELECT name FROM account");
		expect(text()).toContain("SQL is a preview");
		expect(text()).toContain(
			"It comes from FetchXMLToSQL, which isn't documented and can change.",
		);
		const learnMore = [...document.querySelectorAll<HTMLAnchorElement>("a")].find(
			(a) => a.textContent === "Learn more",
		);
		expect(learnMore?.href).toBe(
			"https://learn.microsoft.com/power-apps/developer/data-platform/dataverse-sql-query",
		);
		expect(text()).toContain("From FetchXMLToSQL");
	});
});

describe("a conversion that failed", () => {
	it("says so, with Try again and Copy details, over an empty state", async () => {
		const writeText = vi.fn(async () => {});
		Object.defineProperty(navigator, "clipboard", {
			configurable: true,
			value: { writeText },
		});
		api.fetchXmlToQueryExpression.mockRejectedValueOnce(
			new Error('The link-entity alias "pc" is used twice.'),
		);
		await render(createElement(Code));
		await settle();
		expect(text()).toContain("Couldn't generate C# QueryExpression");
		expect(text()).toContain('The link-entity alias "pc" is used twice.');
		expect(text()).toContain("No code to show");
		expect(text()).toContain("Fix the problem above, then try again.");
		const copy = [
			...document.querySelectorAll<HTMLButtonElement>('[role="toolbar"] button'),
		].find((b) => b.textContent?.trim() === "Copy");
		expect(copy?.getAttribute("aria-disabled")).toBe("true");

		await click(button("Copy details"));
		expect(writeText).toHaveBeenCalledWith('The link-entity alias "pc" is used twice.');
		expect(toasts()).toContain("Copied the error details.");

		await click(button("Try again"));
		await settle();
		expect(api.fetchXmlToQueryExpression).toHaveBeenCalledTimes(2);
		expect(code()).toContain('new QueryExpression("account")');
		expect(text()).not.toContain("No code to show");
	});
});

describe("leaving the Code tab and coming back (#199)", () => {
	it("doesn't ask Dataverse again for the same query", async () => {
		const cache = createConversionCache();
		const panel = () =>
			createElement(CodePanel, {
				fetchXml: XML,
				language: null,
				onLanguageChange: () => {},
				cache,
			});
		await render(panel());
		await settle();
		expect(api.fetchXmlToQueryExpression).toHaveBeenCalledTimes(1);
		cleanup();
		stubOverflowWidths({ container: 3000 });
		await render(panel());
		// The code is there at once, without a second request
		expect(code()).toContain('new QueryExpression("account")');
		await settle();
		expect(api.fetchXmlToQueryExpression).toHaveBeenCalledTimes(1);
	});

	it("shows a result that landed while the tab was closed", async () => {
		const cache = createConversionCache();
		let resolve: (qe: unknown) => void = () => {};
		api.fetchXmlToQueryExpression.mockReturnValue(new Promise((r) => (resolve = r)));
		const panel = () =>
			createElement(CodePanel, {
				fetchXml: XML,
				language: null,
				onLanguageChange: () => {},
				cache,
			});
		await render(panel());
		await settle();
		expect(text()).toContain("Generating…");
		cleanup();
		await act(async () => resolve(QE));
		stubOverflowWidths({ container: 3000 });
		await render(panel());
		await flush(5);
		expect(code()).toContain('new QueryExpression("account")');
		expect(api.fetchXmlToQueryExpression).toHaveBeenCalledTimes(1);
	});
});

describe("Try again", () => {
	it("keeps keyboard focus on the selected language", async () => {
		api.fetchXmlToQueryExpression.mockRejectedValueOnce(new Error("Broken"));
		await render(createElement(Code));
		await settle();
		await click(button("Try again"));
		expect(document.activeElement).toBe(selectedTab());
		expect(labelOf(selectedTab())).toBe("C# QueryExpression");
	});
});

describe("SQL's preview note", () => {
	it("stays while a changed query converts, so it isn't announced again", async () => {
		let setXml: (xml: string) => void = () => {};
		function Sql() {
			const [xml, set] = useState(XML);
			setXml = set;
			return createElement(CodePanel, {
				fetchXml: xml,
				language: "sql",
				onLanguageChange: () => {},
			});
		}
		await render(createElement(Sql));
		await settle();
		const note = () =>
			[...document.querySelectorAll(".fui-MessageBar")].find((m) =>
				m.textContent?.includes("SQL is a preview"),
			);
		const before = note();
		expect(before).toBeTruthy();
		await act(async () =>
			setXml("<fetch><entity name='account'><attribute name='accountid'/></entity></fetch>"),
		);
		await flush(3);
		expect(text()).toContain("Generating…");
		expect(note()).toBe(before);
	});
});
