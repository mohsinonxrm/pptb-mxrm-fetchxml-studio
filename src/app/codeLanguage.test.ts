/**
 * The Code tab's language in the app (#233, decision 2): the one picked stays while you switch tabs,
 * for this session only.
 */
import { afterEach, describe, expect, it, vi } from "vitest";
import { renderApp } from "../shared/testing/app";
import { button, cleanup, click, removeHost } from "../shared/testing/ui";

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
	vi.restoreAllMocks();
});

const FLOW_TIMEOUT = 120000;
const viewTab = (name: string) =>
	[
		...document.querySelectorAll<HTMLElement>(
			'[role="tablist"][aria-label="Query views"] [role="tab"]',
		),
	].find((t) => t.querySelector(".fui-Tab__content")?.textContent?.startsWith(name));
const languageTab = (name: string) =>
	[
		...document.querySelectorAll<HTMLElement>(
			'[role="tablist"][aria-label="Code languages"] [role="tab"]',
		),
	].find((t) => t.querySelector(".fui-Tab__content")?.textContent?.trim() === name);
const selectedLanguage = () =>
	[
		...document.querySelectorAll<HTMLElement>(
			'[role="tablist"][aria-label="Code languages"] [role="tab"][aria-selected="true"]',
		),
	].map((t) => t.querySelector(".fui-Tab__content")?.textContent?.trim())[0];

describe("the Code tab's language", () => {
	it(
		"stays while you switch tabs",
		async () => {
			const app = await renderApp({
				launch: {
					fetchXml:
						"<fetch><entity name='account'><attribute name='name'/></entity></fetch>",
				},
			});
			await app.waitFor(() => !!button("Run", document));
			await click(viewTab("Code"));
			await app.waitFor(() => !!selectedLanguage());
			expect(selectedLanguage()).toBe("C# QueryExpression");

			await click(languageTab("pac CLI"));
			await app.waitFor(() => selectedLanguage() === "pac CLI");
			await click(viewTab("Results"));
			await app.waitFor(() => !selectedLanguage());
			await click(viewTab("Code"));
			await app.waitFor(() => !!selectedLanguage());
			expect(selectedLanguage()).toBe("pac CLI");
		},
		FLOW_TIMEOUT,
	);
});
