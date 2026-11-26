/**
 * Each code editor has a name for screen readers (#279): Monaco's `ariaLabel` option. The stub
 * shows it as the textarea's aria-label, as Monaco does on its own input.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, stubOverflowWidths } from "../../../../shared/testing/ui";
import { FetchXmlEditor } from "./FetchXmlEditor";
import { LayoutXmlViewer } from "./LayoutXmlViewer";
import { CodePanel } from "./CodePanel";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = (props: { value?: string; options?: { ariaLabel?: string } }) =>
		React.createElement("textarea", {
			"data-monaco": "",
			readOnly: true,
			"aria-label": props.options?.ariaLabel,
			value: props.value ?? "",
		});
	return {
		default: Stub,
		Editor: Stub,
		DiffEditor: Stub,
		loader: { config() {}, init: async () => ({}) },
		useMonaco: () => null,
	};
});

afterEach(() => cleanup());

const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";
const editorName = () => document.querySelector("[data-monaco]")?.getAttribute("aria-label");

describe("the code editors' names", () => {
	it("FetchXML: read-only until editor mode, then the editor", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(createElement(FetchXmlEditor, { xml: XML }));
		expect(editorName()).toBe("FetchXML editor, read-only");
		await click(document.querySelector<HTMLInputElement>('input[role="switch"]'));
		expect(editorName()).toBe("FetchXML editor");
	});

	it("LayoutXML", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(createElement(LayoutXmlViewer, { layoutXml: "<grid/>" }));
		expect(editorName()).toBe("LayoutXML, read-only");
	});

	it("Code: names the language", async () => {
		stubOverflowWidths({ container: 2000 });
		await render(
			createElement(CodePanel, {
				fetchXml: XML,
				language: null,
				onLanguageChange: () => {},
			}),
		);
		expect(editorName()).toMatch(/^Code: .+, read-only$/);
	});
});
