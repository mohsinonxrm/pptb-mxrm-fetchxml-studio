/**
 * FetchXML completions, hovers and Format Document are registered once and answer only for FetchXML
 * editors (#194, #186)
 */
import { describe, expect, it } from "vitest";
import type * as Monaco from "monaco-editor/editor";
import { attachFetchXmlLanguageFeatures } from "./fetchxmlIntellisense";

type Provider = Record<string, (...args: unknown[]) => unknown>;

/** A stand-in for Monaco that counts live providers */
function fakeMonaco() {
	const live = {
		completion: new Set<Provider>(),
		hover: new Set<Provider>(),
		formatting: new Set<Provider>(),
	};
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
			registerDocumentFormattingEditProvider: register("formatting"),
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
		getFullModelRange: () => ({
			startLineNumber: 1,
			startColumn: 1,
			endLineNumber: 1,
			endColumn: text.length + 1,
		}),
	}) as unknown as Monaco.editor.ITextModel;

const editor = (m: Monaco.editor.ITextModel) =>
	({ getModel: () => m }) as unknown as Monaco.editor.ICodeEditor;

const position = { lineNumber: 1, column: 3 } as Monaco.Position;

describe("FetchXML language features (#194)", () => {
	it("registers one completion and one hover provider however many editors mount", () => {
		const { monaco, live } = fakeMonaco();
		const handles = [1, 2, 3].map(() =>
			attachFetchXmlLanguageFeatures(monaco, editor(model())),
		);
		expect(live.completion.size).toBe(1);
		expect(live.hover.size).toBe(1);
		expect(live.formatting.size).toBe(1);
		handles.forEach((h) => h.dispose());
		expect(live.formatting.size).toBe(0);
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
		expect(completion.provideCompletionItems(fetchXmlModel, position)).toMatchObject({
			suggestions: expect.any(Array),
		});
		expect(hover.provideHover(fetchXmlModel, position)).not.toBeNull();
		handle.dispose();
	});

	it("formats a FetchXML editor's document (Shift+Alt+F, #186), and nothing else", () => {
		const { monaco, live } = fakeMonaco();
		const messy = '<fetch><entity name="account"><attribute name="name" /></entity></fetch>';
		const fetchXmlModel = model(messy);
		const handle = attachFetchXmlLanguageFeatures(monaco, editor(fetchXmlModel));
		const [formatting] = [...live.formatting];
		const edits = formatting.provideDocumentFormattingEdits(fetchXmlModel) as {
			range: unknown;
			text: string;
		}[];
		expect(edits).toHaveLength(1);
		expect(edits[0].text).toBe(
			[
				"<fetch>",
				'  <entity name="account">',
				'    <attribute name="name" />',
				"  </entity>",
				"</fetch>",
			].join("\n"),
		);
		expect(edits[0].range).toEqual(fetchXmlModel.getFullModelRange());
		// Other XML models, FetchXML that's already formatted, and FetchXML that isn't well-formed
		expect(formatting.provideDocumentFormattingEdits(model(messy))).toEqual([]);
		expect(formatting.provideDocumentFormattingEdits(model(edits[0].text))).toEqual([]);
		handle.dispose();
		const broken = model("<fetch><entity name='account'></fetch>");
		const second = attachFetchXmlLanguageFeatures(monaco, editor(broken));
		expect([...live.formatting][0].provideDocumentFormattingEdits(broken)).toEqual([]);
		const formatted = model(edits[0].text);
		const third = attachFetchXmlLanguageFeatures(monaco, editor(formatted));
		expect([...live.formatting][0].provideDocumentFormattingEdits(formatted)).toEqual([]);
		second.dispose();
		third.dispose();
	});
});
