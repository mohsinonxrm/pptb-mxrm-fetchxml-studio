/**
 * Theme handling against the PPTB 1.2.7 host contract (#67): getCurrentTheme() resolves the
 * "system" setting; settings:updated carries only the changed settings (theme may be "system");
 * the host sends nothing when the OS theme changes, but the webview's colour scheme follows the OS.
 */
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { usePptbContext, type PptbContext } from "./usePptbContext";

type Listener = (event: unknown, payload: { event: string; data: unknown }) => void;

let osDark = false;
let schemeListeners: (() => void)[] = [];
let hostSetting: "light" | "dark" | "system" = "light";
let toolboxListeners: Listener[] = [];
let latest: PptbContext | undefined;
let root: Root;
let container: HTMLDivElement;

const getCurrentTheme = vi.fn(async () => (hostSetting === "system" ? (osDark ? "dark" : "light") : hostSetting));

function Probe() {
	latest = usePptbContext();
	return null;
}

const flush = () => act(async () => { await Promise.resolve(); await Promise.resolve(); });
const mount = async () => {
	await act(async () => root.render(createElement(Probe)));
	await flush();
};
const emitSettings = async (data: unknown) => {
	await act(async () => toolboxListeners.forEach((l) => l({}, { event: "settings:updated", data })));
	await flush();
};
const changeOs = async (dark: boolean) => {
	osDark = dark;
	await act(async () => schemeListeners.forEach((l) => l()));
	await flush();
};

beforeEach(() => {
	(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
	osDark = false;
	schemeListeners = [];
	toolboxListeners = [];
	hostSetting = "light";
	getCurrentTheme.mockClear();
	window.matchMedia = vi.fn(() => ({
		get matches() {
			return osDark;
		},
		addEventListener: (_: string, l: () => void) => schemeListeners.push(l),
		removeEventListener: (_: string, l: () => void) => (schemeListeners = schemeListeners.filter((x) => x !== l)),
	})) as unknown as typeof window.matchMedia;
	(window as unknown as { toolboxAPI: unknown }).toolboxAPI = {
		utils: { getCurrentTheme },
		connections: { getActiveConnection: async () => null },
		events: {
			on: (l: Listener) => toolboxListeners.push(l),
			off: (l: Listener) => (toolboxListeners = toolboxListeners.filter((x) => x !== l)),
		},
	};
	container = document.createElement("div");
	root = createRoot(container);
});

afterEach(async () => {
	await act(async () => root.unmount());
	delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
});

describe("usePptbContext theme (#67)", () => {
	it("starts from the host's resolved theme", async () => {
		hostSetting = "dark";
		await mount();
		expect(latest?.theme).toBe("dark");
	});

	it("asks the host on settings:updated instead of reading the raw setting ('system')", async () => {
		await mount();
		hostSetting = "system";
		osDark = true;
		await emitSettings({ theme: "system" });
		expect(latest?.theme).toBe("dark");
	});

	it("asks the host on settings:updated without a theme key too", async () => {
		await mount();
		hostSetting = "dark";
		await emitSettings({ autoUpdate: true });
		expect(latest?.theme).toBe("dark");
	});

	it("follows the OS while the host setting is System (no host event)", async () => {
		hostSetting = "system";
		await mount();
		expect(latest?.theme).toBe("light");
		await changeOs(true);
		expect(latest?.theme).toBe("dark");
	});

	it("keeps an explicit host theme when the OS changes", async () => {
		hostSetting = "light";
		await mount();
		await changeOs(true);
		expect(latest?.theme).toBe("light");
	});

	it("stops listening when unmounted", async () => {
		await mount();
		await act(async () => root.unmount());
		expect(toolboxListeners).toHaveLength(0);
		expect(schemeListeners).toHaveLength(0);
		root = createRoot(container);
	});

	it("follows the OS when running without the host", async () => {
		delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
		osDark = true;
		await mount();
		expect(latest?.theme).toBe("dark");
		await changeOs(false);
		expect(latest?.theme).toBe("light");
	});
});
