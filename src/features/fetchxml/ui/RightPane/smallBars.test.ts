/**
 * The small bars on the shared overflow bars (#232, PR 3): the FetchXML editor, LayoutXML, Code
 * (language tabs and Copy), Advisories, and the Edit and Add columns drawers. Nothing wraps; what
 * doesn't fit moves into More; unavailable commands keep their reason. jsdom has no layout, so
 * widths come from stubOverflowWidths.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../../../shared/testing/ui";
import { FetchXmlEditor } from "./FetchXmlEditor";
import { LayoutXmlViewer } from "./LayoutXmlViewer";
import { CodePanel } from "./CodePanel";
import { AdvisoriesPanel } from "./AdvisoriesPanel";
import { EditColumnsPanel } from "./EditColumnsPanel";
import { AddColumnsPanel } from "./AddColumnsPanel";
import type { Advisory } from "../../model/advisories";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = (props: { value?: string; defaultValue?: string }) =>
		React.createElement("textarea", {
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

const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";

const toolbar = (name: string) =>
	document.querySelector<HTMLElement>(`[role="toolbar"][aria-label="${name}"]`);
const tabList = (name: string) =>
	document.querySelector<HTMLElement>(`[role="tablist"][aria-label="${name}"]`);
const overflowed = (element: Element | null | undefined) =>
	!!element?.closest("[data-overflowing]");
/** The controls of a bar by name or text, and whether each shows */
const controls = (bar: HTMLElement | null) =>
	[...(bar?.querySelectorAll<HTMLElement>("button") ?? [])].map((b) => ({
		name: b.getAttribute("aria-label") ?? b.textContent?.trim() ?? "",
		element: b,
		shown: !overflowed(b),
	}));
const shownNames = (bar: HTMLElement | null) =>
	controls(bar)
		.filter((c) => c.shown)
		.map((c) => c.name);
const control = (bar: HTMLElement | null, name: string) =>
	controls(bar).find((c) => c.name === name)?.element;
const tab = (list: HTMLElement | null, label: string) =>
	[...(list?.querySelectorAll<HTMLElement>('[role="tab"]') ?? [])].find((t) =>
		t.textContent?.startsWith(label),
	);
const menus = () => [...document.querySelectorAll<HTMLElement>('[role="menu"]')];
const menuTexts = (menu: HTMLElement | undefined) =>
	[...(menu?.querySelectorAll<HTMLElement>('[role^="menuitem"]') ?? [])].map((i) =>
		i.textContent?.trim(),
	);
const description = (element: Element | undefined) => {
	const id = element?.getAttribute("aria-describedby");
	return id ? document.getElementById(id)?.textContent?.trim() : undefined;
};

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("the FetchXML editor's bar", () => {
	const renderEditor = (props: Partial<Parameters<typeof FetchXmlEditor>[0]> = {}) =>
		render(
			createElement(FetchXmlEditor, {
				xml: XML,
				onParseToTree: () => ({ success: true, errors: [], warnings: [] }) as never,
				...props,
			}),
		);
	const bar = () => toolbar("FetchXML commands");
	const editorMode = () => document.querySelector<HTMLInputElement>('input[role="switch"]');

	it("has only Copy in view mode, and the editing commands in editor mode", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderEditor();
		expect(shownNames(bar())).toEqual(["Copy"]);
		await click(editorMode());
		expect(shownNames(bar())).toEqual([
			"Copy",
			"Paste",
			"Load",
			"Reset",
			"Validate",
			"Parse to tree",
		]);
		// The Switch sits beside the bar, never in it
		expect(bar()?.contains(editorMode())).toBe(false);
	});

	it("moves Load, Reset and Paste into More first, then Parse to tree, and keeps Copy longest", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		await renderEditor();
		await click(editorMode());
		// Room for three commands, the divider and More
		await layout.resize({ container: 16 + 3 * 80 + 24 + 40 });
		expect(shownNames(bar())).toEqual(["Copy", "Validate", "Parse to tree", "More commands"]);
		await click(control(bar(), "More commands"));
		expect(menuTexts(menus()[0])).toEqual(["Paste", "Load", "Reset"]);
	});

	it("Paste puts the clipboard's FetchXML in the editor", async () => {
		stubOverflowWidths({ container: 2000 });
		const pasted = "<fetch><entity name='contact'/></fetch>";
		const original = Object.getOwnPropertyDescriptor(navigator, "clipboard");
		Object.defineProperty(navigator, "clipboard", {
			configurable: true,
			value: { readText: async () => pasted },
		});
		try {
			await renderEditor();
			await click(editorMode());
			await click(control(bar(), "Paste"));
			expect(document.querySelector("textarea")?.value).toBe(pasted);
		} finally {
			if (original) Object.defineProperty(navigator, "clipboard", original);
			else delete (navigator as { clipboard?: unknown }).clipboard;
		}
	});

	it("keeps Parse to tree focusable with its reason when there's no tree", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(createElement(FetchXmlEditor, { xml: XML }));
		await click(editorMode());
		const parse = control(bar(), "Parse to tree");
		expect(parse?.getAttribute("aria-disabled")).toBe("true");
		expect(description(parse)).toBe("The query tree isn't available");
	});
});

describe("the LayoutXML bar", () => {
	it("says there's nothing to copy yet", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(createElement(LayoutXmlViewer, { layoutXml: "" }));
		const copy = control(toolbar("LayoutXML commands"), "Copy LayoutXML");
		expect(copy?.getAttribute("aria-disabled")).toBe("true");
		expect(description(copy)).toBe("There's no LayoutXML to copy yet");
		await render(createElement(LayoutXmlViewer, { layoutXml: "<grid/>" }));
		expect(
			control(toolbar("LayoutXML commands"), "Copy LayoutXML")?.getAttribute("aria-disabled"),
		).toBeNull();
	});
});

describe("the Code panel's bar", () => {
	const tabs = () => tabList("Code languages");
	const copy = () => control(toolbar("Code commands"), "Copy code to clipboard");

	it("never wraps the languages: the ones that don't fit move into More", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		await render(createElement(CodePanel, { fetchXml: XML }));
		expect(tab(tabs(), "C# FetchExpression")).toBeTruthy();
		await layout.resize({ container: 3 * 80 + 40 + 10 });
		const more = tabs()?.querySelector<HTMLButtonElement>('button[aria-label^="More tabs"]');
		expect(more).toBeTruthy();
		expect(overflowed(tab(tabs(), "C# FetchExpression"))).toBe(false);
		await click(more);
		const webApi = [
			...(menus()[0]?.querySelectorAll<HTMLElement>('[role="menuitem"]') ?? []),
		].find((i) => i.textContent?.trim() === "Web API");
		await click(webApi);
		expect(tab(tabs(), "Web API")?.getAttribute("aria-selected")).toBe("true");
	});

	it("says why Copy is off: Power Automate's own buttons, or not connected", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(createElement(CodePanel, { fetchXml: XML }));
		expect(copy()?.getAttribute("aria-disabled")).toBeNull();
		await click(tab(tabs(), "Power Automate"));
		expect(copy()?.getAttribute("aria-disabled")).toBe("true");
		expect(description(copy())).toBe("Use the per-field Copy buttons below");
		await click(tab(tabs(), "SQL"));
		expect(copy()?.getAttribute("aria-disabled")).toBe("true");
		expect(description(copy())).toBe("Connect to an environment to generate");
	});
});

describe("the Advisories bar", () => {
	const advisory = (id: string, level: Advisory["level"]): Advisory => ({
		id,
		level,
		title: `Title ${id}`,
		detail: "",
		location: "account",
		doc: { label: "Docs", href: "https://learn.microsoft.com" },
		source: "fxs",
	});
	const renderPanel = (extra: Partial<Parameters<typeof AdvisoriesPanel>[0]> = {}) =>
		render(
			createElement(AdvisoriesPanel, {
				advisories: [advisory("a", "error"), advisory("b", "advice")],
				noQuery: false,
				editorInvalid: false,
				inTree: true,
				checking: false,
				stale: false,
				added: 0,
				onDismissCheckError: () => {},
				onCheck: () => {},
				...extra,
			}),
		);
	const bar = () => toolbar("Advisories");
	const radio = (label: string) =>
		[...(bar()?.querySelectorAll<HTMLElement>('[role="radio"]') ?? [])].find((r) =>
			r.textContent?.startsWith(label),
		);

	it("filters with one radio group, and Check stays when narrow", async () => {
		const layout = stubOverflowWidths({ container: 2000 });
		await renderPanel();
		expect(radio("All (2)")?.getAttribute("aria-checked")).toBe("true");
		await click(radio("Will be rejected (1)"));
		expect(radio("Will be rejected (1)")?.getAttribute("aria-checked")).toBe("true");
		expect(document.body.textContent).toContain("Title a");
		expect(document.body.textContent).not.toContain("Title b");

		// Room for Check, one filter, the divider and More: the chosen filter stays on the bar
		await layout.resize({ container: 2 * 80 + 2 * 24 + 40 + 16 });
		expect(overflowed(control(bar(), "Check with Dataverse"))).toBe(false);
		expect(overflowed(radio("Will be rejected (1)"))).toBe(false);
		expect(overflowed(radio("All (2)"))).toBe(true);

		// Choosing a filter from More filters the list
		await click(control(bar(), "More commands"));
		const advice = [
			...(menus()[0]?.querySelectorAll<HTMLElement>('[role="menuitemradio"]') ?? []),
		].find((i) => i.textContent?.trim() === "Performance advice (1)");
		await click(advice);
		expect(document.body.textContent).toContain("Title b");
		expect(document.body.textContent).not.toContain("Title a");
	});

	it("names each filter 'Show …', and keeps the status as a line of its own", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderPanel();
		expect(radio("All (2)")?.getAttribute("aria-label")).toBe("Show all (2)");
		expect(radio("Will be rejected (1)")?.getAttribute("aria-label")).toBe(
			"Show will be rejected (1)",
		);
		const status = document.querySelector('[aria-live="polite"]');
		expect(status).toBeTruthy();
		expect(bar()?.contains(status)).toBe(false);
	});

	it("keeps Check focusable with its reason", async () => {
		stubOverflowWidths({ container: 2000 });
		await renderPanel({ noQuery: true });
		const check = control(bar(), "Check with Dataverse");
		expect(check?.getAttribute("aria-disabled")).toBe("true");
		expect(description(check)).toBe("Choose an entity first");
		await renderPanel({ checkDisabledReason: "Connect to an environment to check" });
		expect(description(control(bar(), "Check with Dataverse"))).toBe(
			"Connect to an environment to check",
		);
	});
});

describe("the column drawers", () => {
	it("Edit columns: Add columns keeps its reason, on the bar and as text", async () => {
		stubOverflowWidths({ container: 2000 });
		const reason = "Aggregate queries can't add columns";
		await render(
			createElement(EditColumnsPanel, {
				open: true,
				columns: [],
				onClose: () => {},
				onApply: () => {},
				onAddColumns: () => {},
				onResetToDefault: () => {},
				addColumnsDisabledReason: reason,
			}),
		);
		const add = control(toolbar("Column commands"), "Add columns");
		expect(add?.getAttribute("aria-disabled")).toBe("true");
		expect(description(add)).toBe(reason);
		expect(control(toolbar("Column commands"), "Reset to default")).toBeTruthy();
	});

	it("Add columns: the entity and Related tabs switch", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(
			createElement(AddColumnsPanel, {
				open: true,
				entityDisplayName: "Account",
				onClose: () => {},
				onApply: () => {},
			}),
		);
		const list = tabList("Columns from");
		expect(tab(list, "Account")?.getAttribute("aria-selected")).toBe("true");
		await click(tab(list, "Related"));
		expect(tab(tabList("Columns from"), "Related")?.getAttribute("aria-selected")).toBe("true");
	});
});
