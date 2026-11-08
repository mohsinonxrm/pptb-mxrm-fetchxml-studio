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
import { createFakeDataverse, ORG_URL, sampleTables, type FakeDataverseOptions, type FakeTable } from "./fakeDataverse";
import { flush, render } from "./ui";
import { resetDataverseSessionCache } from "../../features/fetchxml/api/sessionCache";
import { clearMetadataCache } from "../../features/fetchxml/api/dataverseMetadata";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { clearMetadataLoadErrors } from "../../features/fetchxml/ui/LeftPane/PropertiesPanel/metadataLoadErrors";

export interface AppHarnessOptions extends FakeDataverseOptions {
	tables?: FakeTable[];
	/** What a launching tool passes (tool-to-tool prefill): { fetchXml } or { viewRef } */
	launch?: Record<string, unknown> | null;
	connection?: { id: string; name: string; url: string; environment: string };
}

export interface AppHarness {
	dataverse: ReturnType<typeof createFakeDataverse>;
	notifications: { title?: string; body?: string; type?: string }[];
	/** The status line under the results ("Rows: 3 (more available) | Executed in 2ms"), or "" */
	status: () => string;
	/** Close holds taken (toolboxAPI.preventClose) and released: equal once nothing is running */
	closeHolds: { prevented: number; released: number };
	/**
	 * Fails if the run left something behind: a Dataverse request the fake doesn't know, a page sent
	 * with a missing or wrong paging cookie, or a close hold that was never released
	 */
	expectClean: () => void;
	/** Waits until the predicate holds, flushing React in between; fails after `ms` */
	waitFor: (predicate: () => boolean, ms?: number) => Promise<void>;
}

export async function renderApp(options: AppHarnessOptions = {}): Promise<AppHarness> {
	// Module-level caches outlive a render: forget the previous test's Dataverse, as a connection
	// change does (AppShell), so each test starts from a fresh session
	resetDataverseSessionCache();
	clearMetadataCache();
	metadataCache.clear();
	clearMetadataLoadErrors();
	const dataverse = createFakeDataverse(options.tables ?? sampleTables(), options);
	const connection = options.connection ?? { id: "c1", name: "Test environment", url: ORG_URL, environment: "Dev" };
	const notifications: AppHarness["notifications"] = [];
	const settings = new Map<string, unknown>();
	const closeHolds = { prevented: 0, released: 0 };
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
		// PPTB >= 1.2.6: long operations hold the tab open (host.ts holdClose)
		preventClose: async () => {
			closeHolds.prevented++;
		},
		releasePreventClose: async () => {
			closeHolds.released++;
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

	// The results info bar's first span (ResultsGrid). Document order puts the outermost span first.
	const status = () => {
		const span = [...document.querySelectorAll("span")].find((e) => e.textContent?.startsWith("Rows:"));
		return (span?.textContent ?? "").replace(/\s+/g, " ").trim();
	};
	// Generous by default: a wait returns as soon as its predicate holds, and a busy CI runner or a
	// parallel full suite can take several seconds per page (seen: flakes at 4 s)
	const waitFor = async (predicate: () => boolean, ms = 15000) => {
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
	const expectClean = () => {
		const problems = [
			...dataverse.unknown.map((q) => `unknown request: ${q}`),
			...dataverse.pagingErrors,
			...(closeHolds.prevented !== closeHolds.released ? [`close holds: ${closeHolds.prevented} taken, ${closeHolds.released} released`] : []),
		];
		if (problems.length) throw new Error(`The run left problems:\n${problems.join("\n")}`);
	};
	return { dataverse, notifications, status, closeHolds, expectClean, waitFor };
}
