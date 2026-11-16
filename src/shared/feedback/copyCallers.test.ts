/**
 * Every Copy in the tool says so in the app's Toaster (#277): "Copied {what}", or "Couldn't copy to
 * the clipboard", and the button keeps its own label (no "Copied!" swap). The FetchXML editor's
 * short confirmations are toasts too; its bars stay for problems in the XML.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	button,
	cleanup,
	click,
	flush,
	installHost,
	removeHost,
	render,
	toasts,
} from "../testing/ui";
import { CopyButton } from "../components/ErrorFallbacks";
import { LayoutXmlViewer } from "../../features/fetchxml/ui/RightPane/LayoutXmlViewer";
import { PowerAutomatePane } from "../../features/fetchxml/ui/RightPane/PowerAutomatePane";
import { CodePanel } from "../../features/fetchxml/ui/RightPane/CodePanel";
import { FetchXmlEditor } from "../../features/fetchxml/ui/RightPane/FetchXmlEditor";
import { BulkDeleteDialog } from "../../features/fetchxml/ui/Dialogs/BulkDeleteDialog";
import { parseFetchXml } from "../../features/fetchxml/model/fetchxmlParser";

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

/** What the ToolBox clipboard received */
let copied: string[];

beforeEach(() => {
	installHost();
	copied = [];
	const utils = (window as unknown as { toolboxAPI: { utils: Record<string, unknown> } })
		.toolboxAPI.utils;
	utils.copyToClipboard = async (text: string) => {
		copied.push(text);
	};
});

/** jsdom's own navigator.clipboard (none), put back after a test replaces it */
const originalClipboard = Object.getOwnPropertyDescriptor(navigator, "clipboard");
const setBrowserClipboard = (clipboard: Partial<Clipboard>) =>
	Object.defineProperty(navigator, "clipboard", { configurable: true, value: clipboard });

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
	if (originalClipboard) Object.defineProperty(navigator, "clipboard", originalClipboard);
	else delete (navigator as { clipboard?: unknown }).clipboard;
});

/** Makes both clipboards fail: the ToolBox's, then the browser's fallback */
const breakClipboard = () => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
	const utils = (window as unknown as { toolboxAPI: { utils: Record<string, unknown> } })
		.toolboxAPI.utils;
	utils.copyToClipboard = async () => {
		throw new Error("denied");
	};
	setBrowserClipboard({ writeText: async () => Promise.reject(new Error("denied")) });
};

describe("copy buttons", () => {
	it("the error fallback's Copy details names what it copied, and keeps its label", async () => {
		await render(
			createElement(CopyButton, {
				text: "Error: boom",
				label: "Copy details",
				what: "error details",
			}),
		);
		await click(button("Copy details"));
		expect(copied).toEqual(["Error: boom"]);
		expect(toasts()).toEqual(["Copied error details."]);
		expect(button("Copy details")).toBeDefined();
	});

	it("says when the clipboard can't be written", async () => {
		breakClipboard();
		await render(
			createElement(CopyButton, { text: "x", label: "Copy FetchXML", what: "FetchXML" }),
		);
		await click(button("Copy FetchXML"));
		expect(toasts()).toEqual([expect.stringMatching(/^Couldn't copy to the clipboard\./)]);
	});

	it("LayoutXML", async () => {
		await render(createElement(LayoutXmlViewer, { layoutXml: "<grid><row/></grid>" }));
		await click(document.querySelector('[aria-label="Copy LayoutXML"]'));
		expect(copied).toHaveLength(1);
		expect(toasts()).toEqual(["Copied LayoutXML."]);
	});

	it("Power Automate: a field by its label, and all fields", async () => {
		await render(
			createElement(PowerAutomatePane, {
				spec: {
					actionName: "List rows",
					connector: "Microsoft Dataverse connector",
					fields: [
						{ label: "Table name", value: "Accounts" },
						{ label: "Fetch Xml Query", value: "<fetch/>" },
						{ label: "Filter rows", value: "" },
					],
				},
			}),
		);
		const copyButtons = [...document.querySelectorAll("button")].filter(
			(b) => b.textContent?.trim() === "Copy",
		);
		await click(copyButtons[0]);
		expect(copied).toEqual(["Accounts"]);
		expect(toasts()).toEqual(["Copied Table name."]);
		expect(copyButtons[0].textContent?.trim()).toBe("Copy");

		await click(button("Copy all"));
		expect(copied).toHaveLength(2);
		expect(toasts()).toEqual(["Copied Table name.", "Copied all fields."]);
	});

	it("Code: names the language, and the button still reads Copy", async () => {
		await render(
			createElement(CodePanel, {
				fetchXml: "<fetch><entity name='account'><attribute name='name'/></entity></fetch>",
			}),
		);
		await flush(5);
		await click(button("Copy"));
		expect(copied).toHaveLength(1);
		expect(copied[0]).toContain("account");
		expect(toasts()).toEqual(["Copied C# FetchExpression."]);
		expect(button("Copy")).toBeDefined();
	});

	it("Bulk delete: the job URL", async () => {
		await render(
			createElement(BulkDeleteDialog, {
				open: true,
				recordCount: 3,
				entityDisplayName: "Account",
				entityLogicalName: "account",
				onClose: () => {},
				onConfirm: async () => ({
					asyncOperationId: "job-1",
					jobUrl: "https://test.crm.dynamics.com/job-1",
				}),
			}),
		);
		await click(button("Submit bulk delete job") ?? button(/^Submit/));
		await flush(5);
		await click(button("Copy URL"));
		expect(copied).toEqual(["https://test.crm.dynamics.com/job-1"]);
		expect(toasts()).toEqual(["Copied the job URL."]);
		expect(button("Copy URL")).toBeDefined();
	});
});

describe("the FetchXML editor", () => {
	const editorMode = () =>
		click(document.querySelector('input[role="switch"]') as HTMLInputElement | null);
	const bars = () =>
		[...document.querySelectorAll(".fui-MessageBar")].map((b) =>
			(b.textContent ?? "").replace(/\s+/g, " ").trim(),
		);

	it("Copy is a toast, not a bar", async () => {
		await render(
			createElement(FetchXmlEditor, { xml: "<fetch><entity name='account'/></fetch>" }),
		);
		await click(button("Copy"));
		expect(copied).toEqual(["<fetch><entity name='account'/></fetch>"]);
		expect(toasts()).toEqual(["Copied FetchXML."]);
		expect(bars()).toEqual([]);
	});

	it("Valid is a toast; a syntax error stays a bar in the tab", async () => {
		await render(
			createElement(FetchXmlEditor, {
				xml: "<fetch><entity name='account'><attribute name='name'/></entity></fetch>",
			}),
		);
		await editorMode();
		await click(button("Validate"));
		expect(toasts()).toEqual(["Valid FetchXML. Dataverse should accept it.."]);
		expect(bars()).toEqual([]);

		cleanup();
		await render(createElement(FetchXmlEditor, { xml: "<fetch><entity name='account'>" }));
		await editorMode();
		await click(button("Validate"));
		expect(bars()).toEqual([expect.stringMatching(/^Syntax error/)]);
		expect(toasts()).toEqual([]);
	});

	it("Reset is a toast", async () => {
		await render(
			createElement(FetchXmlEditor, { xml: "<fetch><entity name='account'/></fetch>" }),
		);
		await editorMode();
		await click(button("Reset"));
		expect(toasts()).toEqual(["Reverted to the FetchXML from the tree."]);
		expect(bars()).toEqual([]);
	});

	it("Paste is a toast, and a clipboard it can't read says so", async () => {
		const onEditorStateChange = vi.fn();
		setBrowserClipboard({ readText: async () => "<fetch><entity name='contact'/></fetch>" });
		await render(
			createElement(FetchXmlEditor, {
				xml: "<fetch><entity name='account'/></fetch>",
				onEditorStateChange,
			}),
		);
		await editorMode();
		await click(button("Paste"));
		expect(onEditorStateChange).toHaveBeenLastCalledWith(
			true,
			"<fetch><entity name='contact'/></fetch>",
		);
		expect(toasts()).toEqual(["Pasted from the clipboard."]);

		vi.spyOn(console, "error").mockImplementation(() => {});
		setBrowserClipboard({ readText: async () => Promise.reject(new Error("denied")) });
		await click(button("Paste"));
		expect(toasts()[1]).toMatch(/^Couldn't read the clipboard\./);
		expect(bars()).toEqual([]);
	});

	it("Parse to tree: a toast, and editor mode ends", async () => {
		const onEditorStateChange = vi.fn();
		await render(
			createElement(FetchXmlEditor, {
				xml: "<fetch><entity name='account'><attribute name='name'/></entity></fetch>",
				onParseToTree: (xml: string) => parseFetchXml(xml),
				onEditorStateChange,
			}),
		);
		await editorMode();
		await click(button("Parse to tree"));
		expect(toasts()).toEqual(["Loaded into the tree."]);
		expect(onEditorStateChange).toHaveBeenLastCalledWith(false, "");
		expect(bars()).toEqual([]);
	});

	it("Editor updated: a toast when the tree replaces the editor's FetchXML", async () => {
		const xml = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";
		await render(createElement(FetchXmlEditor, { xml, replaceRequest: null }));
		await editorMode();
		await render(
			createElement(FetchXmlEditor, {
				xml,
				replaceRequest: { xml: "<fetch><entity name='contact'/></fetch>", id: 1 },
			}),
		);
		expect(toasts()).toEqual([
			"Editor updated. The editor now shows the FetchXML that was loaded into the tree, with the changes you confirmed..",
		]);
		expect(bars()).toEqual([]);
	});
});
