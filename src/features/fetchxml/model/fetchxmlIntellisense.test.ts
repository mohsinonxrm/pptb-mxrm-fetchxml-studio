/**
 * FetchXML completions and hovers are registered once and answer only for FetchXML editors (#194)
 */
import { describe, expect, it } from "vitest";
import type * as Monaco from "monaco-editor/editor";
import { attachFetchXmlLanguageFeatures } from "./fetchxmlIntellisense";

type Provider = Record<string, (...args: unknown[]) => unknown>;

/** A stand-in for Monaco that counts live providers */
function fakeMonaco() {
	const live = { completion: new Set<Provider>(), hover: new Set<Provider>() };
	const register = (kind: keyof typeof live) => (_language: string, provider: Provider) => {
		live[kind].add(provider);
		return { dispose: () => live[kind].delete(provider) };
	};
	const monaco = {
		languages: {
			// Enum members the providers read (CompletionItemKind.Property, …): any number will do
			CompletionItemKind: new Proxy({}, { get: () => 1 }),
			CompletionItemInsertTextRule: new Proxy({}, { get: () => 4 }),
			registerCompletionItemProvider: register("completion"),
			registerHoverProvider: register("hover"),
		},
	} as unknown as typeof Monaco;
	return { monaco, live };
}

/** A text model with one line of FetchXML */
const model = (text = '<fetch><entity name="account" /></fetch>') =>
	({
		getWordUntilPosition: () => ({ startColumn: 1, endColumn: 1, word: "" }),
		getWordAtPosition: () => ({ startColumn: 2, endColumn: 7, word: "fetch" }),
		getLineContent: () => text,
		getValueInRange: () => text,
		getLineCount: () => 1,
		getValue: () => text,
	}) as unknown as Monaco.editor.ITextModel;

const editor = (m: Monaco.editor.ITextModel) => ({ getModel: () => m }) as unknown as Monaco.editor.ICodeEditor;

const position = { lineNumber: 1, column: 3 } as Monaco.Position;

describe("FetchXML language features (#194)", () => {
	it("registers one completion and one hover provider however many editors mount", () => {
		const { monaco, live } = fakeMonaco();
		const handles = [1, 2, 3].map(() => attachFetchXmlLanguageFeatures(monaco, editor(model())));
		expect(live.completion.size).toBe(1);
		expect(live.hover.size).toBe(1);
		handles.forEach((h) => h.dispose());
	});

	it("removes the providers when the last editor goes, and registers again for the next one", () => {
		const { monaco, live } = fakeMonaco();
		const first = attachFetchXmlLanguageFeatures(monaco, editor(model()));
		const second = attachFetchXmlLanguageFeatures(monaco, editor(model()));
		first.dispose();
		expect(live.completion.size).toBe(1);
		second.dispose();
		expect(live.completion.size).toBe(0);
		expect(live.hover.size).toBe(0);

		// A connection change remounts the editor
		const remounted = attachFetchXmlLanguageFeatures(monaco, editor(model()));
		expect(live.completion.size).toBe(1);
		remounted.dispose();
	});

	it("ignores a second dispose of the same handle", () => {
		const { monaco, live } = fakeMonaco();
		const first = attachFetchXmlLanguageFeatures(monaco, editor(model()));
		const second = attachFetchXmlLanguageFeatures(monaco, editor(model()));
		first.dispose();
		first.dispose();
		expect(live.completion.size).toBe(1);
		second.dispose();
	});

	it("answers for the FetchXML editor's model only, not other XML models such as LayoutXML", () => {
		const { monaco, live } = fakeMonaco();
		const fetchXmlModel = model();
		const layoutXmlModel = model('<grid name="resultset"><row name="result" /></grid>');
		const handle = attachFetchXmlLanguageFeatures(monaco, editor(fetchXmlModel));
		const [completion] = [...live.completion];
		const [hover] = [...live.hover];

		expect(completion.provideCompletionItems(layoutXmlModel, position)).toBeUndefined();
		expect(hover.provideHover(layoutXmlModel, position)).toBeNull();
		expect(completion.provideCompletionItems(fetchXmlModel, position)).toMatchObject({ suggestions: expect.any(Array) });
		expect(hover.provideHover(fetchXmlModel, position)).not.toBeNull();
		handle.dispose();
	});
});
