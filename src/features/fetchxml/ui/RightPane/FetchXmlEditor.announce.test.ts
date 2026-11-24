/**
 * The FetchXML editor's bars for screen readers (#186 PR 3, #277 note): a new bar announces itself
 * (Fluent, through the app's announcer); a bar updated in place doesn't remount, so the editor
 * announces it: errors and warnings assertively, info politely.
 */
import { act, createElement, createRef } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../../../shared/testing/ui";
import { AnnouncerProvider } from "../../../../shared/feedback/Announcer";
import { FLUSH_MS } from "../../../../shared/feedback/useAnnouncer";
import { FetchXmlEditor, type FetchXmlEditorHandle } from "./FetchXmlEditor";
import type { EditorProblem } from "./editorProblems";

vi.mock("@monaco-editor/react", async () => {
	const React = await import("react");
	const Stub = (props: { value?: string }) =>
		React.createElement("textarea", { readOnly: true, value: props.value ?? "" });
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
const assertive = () =>
	document.querySelector('[data-fxs-announcer="assertive"]')?.textContent ?? "";
const polite = () => document.querySelector('[data-fxs-announcer="polite"]')?.textContent ?? "";
const flushed = () => act(() => new Promise<void>((resolve) => setTimeout(resolve, FLUSH_MS + 40)));
const syntax = (body: string): EditorProblem => ({
	key: "error",
	intent: "error",
	title: "Syntax error",
	body,
});

describe("the editor's bars and the announcer", () => {
	it("announces a bar updated in place, assertively for an error", async () => {
		const ref = createRef<FetchXmlEditorHandle>();
		await render(
			createElement(
				AnnouncerProvider,
				null,
				createElement(FetchXmlEditor, { xml: XML, ref }),
			),
		);
		await act(async () => ref.current?.report(["error"], [syntax("Line 3: missing quote")]));
		await flushed();
		// The new bar announced itself
		expect(assertive()).toContain("Line 3: missing quote");
		// The same bar, new text: the editor announces it
		await act(async () => ref.current?.report(["error"], [syntax("Line 9: unclosed tag")]));
		await flushed();
		expect(assertive()).toBe("Syntax error. Line 9: unclosed tag");
	});

	it("announces an info bar updated in place politely, with its findings", async () => {
		const ref = createRef<FetchXmlEditorHandle>();
		await render(
			createElement(
				AnnouncerProvider,
				null,
				createElement(FetchXmlEditor, { xml: XML, ref }),
			),
		);
		const notes = (message: string): EditorProblem => ({
			key: "notes",
			intent: "info",
			title: "Notes",
			body: [{ element: "order", message }],
		});
		await act(async () => ref.current?.report(["notes"], [notes("First note")]));
		await flushed();
		await act(async () => ref.current?.report(["notes"], [notes("Second note")]));
		await flushed();
		expect(polite()).toBe("Notes. <order> Second note");
	});
});
