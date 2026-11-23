/**
 * The FetchXML editor's Monaco hooks (#186): Ctrl+Enter is Monaco's own "Run query" action, in the
 * text only, and runs through the editor sync's Run (which waits while Run has a reason); the editor
 * is marked so the app's key handler leaves Ctrl+Enter to it; Format runs Monaco's Format Document
 * (one undo step, cursor kept) once the formatter has checked the FetchXML.
 */
import { createElement, type ReactNode } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render, toasts } from "../../../../shared/testing/ui";
import { RUNS_ITSELF_ATTRIBUTE } from "../../../../app/shortcuts/registry";
import { EditorSyncContext, type EditorSyncValue } from "./editorSync";
import { FetchXmlEditor } from "./FetchXmlEditor";

const fake = vi.hoisted(() => {
	const state = {
		text: "",
		actions: [] as {
			id: string;
			label: string;
			keybindings?: number[];
			keybindingContext?: string;
			run: () => void;
		}[],
		formatRun: null as null | (() => Promise<void>),
		dom: null as HTMLElement | null,
	};
	return state;
});

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	function Stub(props: { value?: string; onMount?: (editor: unknown, monaco: unknown) => void }) {
		fake.text = props.value ?? "";
		const mounted = React.useRef(false);
		const ref = React.useRef<HTMLDivElement>(null);
		React.useEffect(() => {
			if (mounted.current || !props.onMount) return;
			mounted.current = true;
			fake.dom = ref.current;
			const model = {
				getValue: () => fake.text,
				getFullModelRange: () => ({
					startLineNumber: 1,
					startColumn: 1,
					endLineNumber: 1,
					endColumn: 1,
				}),
			};
			const editor = {
				getModel: () => model,
				getDomNode: () => ref.current,
				onDidDispose: () => ({ dispose() {} }),
				addAction: (action: (typeof fake.actions)[number]) => {
					fake.actions.push(action);
					return { dispose() {} };
				},
				getAction: (id: string) =>
					id === "editor.action.formatDocument" && fake.formatRun
						? { run: fake.formatRun }
						: null,
			};
			const register = () => ({ dispose() {} });
			const monaco = {
				KeyMod: { CtrlCmd: 2048 },
				KeyCode: { Enter: 3 },
				languages: {
					CompletionItemKind: new Proxy({}, { get: () => 1 }),
					CompletionItemInsertTextRule: new Proxy({}, { get: () => 4 }),
					registerCompletionItemProvider: register,
					registerHoverProvider: register,
					registerDocumentFormattingEditProvider: register,
				},
			};
			props.onMount(editor, monaco);
		});
		return React.createElement(
			"div",
			{ ref, className: "monaco-editor" },
			React.createElement("textarea", { readOnly: true, value: props.value ?? "" }),
		);
	}
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
	fake.actions = [];
	fake.formatRun = null;
	fake.dom = null;
});

const XML = "<fetch><entity name='account'><attribute name='name'/></entity></fetch>";
const withSync = (sync: Partial<EditorSyncValue>, child: ReactNode) =>
	createElement(
		EditorSyncContext.Provider,
		{
			value: {
				sync: async () => ({ status: "unchanged" as const }),
				editorHasContent: true,
				...sync,
			},
		},
		child,
	);
const editorMode = () => document.querySelector<HTMLInputElement>('input[role="switch"]');
const formatButton = () =>
	[...document.querySelectorAll<HTMLButtonElement>('[role="toolbar"] button')].find(
		(b) => b.textContent?.trim() === "Format",
	);

describe("the FetchXML editor's Monaco hooks (#186)", () => {
	it("registers Run query on Ctrl+Enter, in the text only, through the sync's Run", async () => {
		const runFromEditor = vi.fn();
		await render(withSync({ runFromEditor }, createElement(FetchXmlEditor, { xml: XML })));
		const run = fake.actions.find((a) => a.id === "fxs.runQuery");
		expect(run).toMatchObject({
			label: "Run query",
			keybindings: [2048 | 3],
			keybindingContext: "editorTextFocus",
		});
		run?.run();
		expect(runFromEditor).toHaveBeenCalledOnce();
		expect(fake.dom?.hasAttribute(RUNS_ITSELF_ATTRIBUTE)).toBe(true);
	});

	it("Format runs Monaco's Format Document when the FetchXML can be formatted", async () => {
		fake.formatRun = vi.fn(async () => {});
		await render(withSync({}, createElement(FetchXmlEditor, { xml: XML })));
		await click(editorMode());
		await click(formatButton());
		expect(fake.formatRun).toHaveBeenCalledOnce();
	});

	it("Format says why, and leaves Monaco alone, when the FetchXML isn't well-formed", async () => {
		fake.formatRun = vi.fn(async () => {});
		await render(
			withSync(
				{},
				createElement(FetchXmlEditor, { xml: "<fetch><entity name='a'></fetch>" }),
			),
		);
		await click(editorMode());
		await click(formatButton());
		expect(fake.formatRun).not.toHaveBeenCalled();
		expect(toasts().join(" ")).toContain("Can't format: the FetchXML isn't well-formed");
	});

	it("Format does nothing when the FetchXML is already formatted", async () => {
		fake.formatRun = vi.fn(async () => {});
		const formatted = ["<fetch>", "  <entity name='account' />", "</fetch>"].join("\n");
		await render(withSync({}, createElement(FetchXmlEditor, { xml: formatted })));
		await click(editorMode());
		await click(formatButton());
		expect(fake.formatRun).not.toHaveBeenCalled();
	});
});
