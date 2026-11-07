/**
 * App-level test setup (#189, #301): the whole tool rendered against the in-memory Dataverse
 * (fakeDataverse.ts) and a stubbed ToolBox host, launched the way another tool would launch it.
 * Test files must also mock "@monaco-editor/react" (see appFlows.test.ts): vi.mock only applies in
 * the test file itself.
 *
 * Only test files import this module.
 */
import { createElement } from "react";
import { act } from "react";
import { createFakeDataverse, sampleTables, type FakeDataverseOptions, type FakeTable } from "./fakeDataverse";
import { flush, render } from "./ui";

export interface AppHarnessOptions extends FakeDataverseOptions {
	tables?: FakeTable[];
	/** What a launching tool passes (tool-to-tool prefill): { fetchXml } or { viewRef } */
	launch?: Record<string, unknown> | null;
	connection?: { id: string; name: string; url: string; environment: string };
}

export interface AppHarness {
	dataverse: ReturnType<typeof createFakeDataverse>;
	notifications: { title?: string; body?: string; type?: string }[];
	/** The text of the status line under the results (e.g. "Rows: 3 | Executed in 2ms"), or "" */
	status: () => string;
	/** Waits until the predicate holds, flushing React in between; fails after `ms` */
	waitFor: (predicate: () => boolean, ms?: number) => Promise<void>;
}

export async function renderApp(options: AppHarnessOptions = {}): Promise<AppHarness> {
	const dataverse = createFakeDataverse(options.tables ?? sampleTables(), options);
	const connection = options.connection ?? { id: "c1", name: "Test environment", url: "https://test.crm.dynamics.com/", environment: "Dev" };
	const notifications: AppHarness["notifications"] = [];
	const settings = new Map<string, unknown>();
	(window as unknown as { dataverseAPI: unknown }).dataverseAPI = dataverse.api;
	(window as unknown as { toolboxAPI: unknown }).toolboxAPI = {
		connections: { getConnection: async () => connection, getActiveConnection: async () => connection },
		utils: {
			showNotification: async (n: AppHarness["notifications"][number]) => {
				notifications.push(n);
			},
			getCurrentTheme: async () => "light",
			copyToClipboard: async () => undefined,
		},
		events: { on: () => undefined, off: () => undefined },
		settings: {
			get: async (key: string) => settings.get(key),
			set: async (key: string, value: unknown) => {
				settings.set(key, value);
			},
			getAll: async () => Object.fromEntries(settings),
		},
		getToolContext: async () => ({ toolId: "local-test", instanceId: "test-instance" }),
		invocation: options.launch === null ? undefined : { getLaunchContext: async () => options.launch ?? null },
	};
	// jsdom has neither; the tree scrolls its selected node into view, Fluent reads media queries
	Element.prototype.scrollIntoView = function () {};
	window.matchMedia = ((query: string) => ({
		matches: false,
		media: query,
		addEventListener() {},
		removeEventListener() {},
		addListener() {},
		removeListener() {},
	})) as unknown as typeof window.matchMedia;

	const { AppShell } = await import("../../app/AppShell");
	await render(createElement(AppShell));

	const status = () => {
		const text = document.body.textContent ?? "";
		return /Rows: [^|]*\|?[^A-Z]*/.exec(text)?.[0]?.trim() ?? "";
	};
	const waitFor = async (predicate: () => boolean, ms = 4000) => {
		const end = Date.now() + ms;
		while (!predicate()) {
			if (Date.now() > end) throw new Error(`waitFor timed out. Status: "${status()}"`);
			await act(async () => {
				await new Promise((resolve) => setTimeout(resolve, 20));
			});
		}
		await flush();
	};
	await waitFor(() => dataverse.calls.some((c) => c.startsWith("execute WhoAmI")));
	return { dataverse, notifications, status, waitFor };
}
