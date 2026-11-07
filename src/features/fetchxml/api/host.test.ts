import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { canExecuteBatch, connectionGuard, copyText, getPrimaryConnection, holdClose, openUrl, whileHoldingClose } from "./host";

// #70: host APIs are feature-detected, so older hosts (and the dev server) degrade gracefully

const setToolbox = (api: unknown) => ((window as unknown as { toolboxAPI?: unknown }).toolboxAPI = api);
const connection = (id: string, url: string) => ({ id, url, name: id, environment: "Dev" as const });

let clipboardWrite: ReturnType<typeof vi.fn>;

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
	clipboardWrite = vi.fn().mockResolvedValue(undefined);
	Object.defineProperty(navigator, "clipboard", { value: { writeText: clipboardWrite }, configurable: true });
});

afterEach(() => {
	vi.restoreAllMocks();
	delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
	delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
});

describe("getPrimaryConnection", () => {
	it("uses the slot-aware getConnection(0) when the host has it (PPTB ≥ 1.2.7)", async () => {
		const getConnection = vi.fn(async () => connection("a", "https://a.crm.dynamics.com"));
		const getActiveConnection = vi.fn();
		setToolbox({ connections: { getConnection, getActiveConnection } });
		expect((await getPrimaryConnection())?.id).toBe("a");
		expect(getConnection).toHaveBeenCalledWith(0);
		expect(getActiveConnection).not.toHaveBeenCalled();
	});

	it("falls back to getActiveConnection, and to null without a host", async () => {
		setToolbox({ connections: { getActiveConnection: async () => connection("b", "https://b") } });
		expect((await getPrimaryConnection())?.id).toBe("b");
		setToolbox(undefined);
		expect(await getPrimaryConnection()).toBeNull();
	});
});

describe("openUrl", () => {
	it("opens in the connection's browser when the host can", async () => {
		const openInConnectionBrowser = vi.fn().mockResolvedValue(undefined);
		const open = vi.spyOn(window, "open").mockImplementation(() => null);
		setToolbox({ utils: { openInConnectionBrowser } });
		await openUrl("https://org.crm.dynamics.com/main.aspx");
		expect(openInConnectionBrowser).toHaveBeenCalledWith("https://org.crm.dynamics.com/main.aspx");
		expect(open).not.toHaveBeenCalled();
	});

	it("falls back to window.open", async () => {
		const open = vi.spyOn(window, "open").mockImplementation(() => null);
		setToolbox({ utils: { openInConnectionBrowser: vi.fn().mockRejectedValue(new Error("no profile")) } });
		await openUrl("https://x");
		expect(open).toHaveBeenCalledWith("https://x", "_blank", "noopener");
	});
});

describe("copyText", () => {
	it("uses the ToolBox clipboard first", async () => {
		const copyToClipboard = vi.fn().mockResolvedValue(undefined);
		setToolbox({ utils: { copyToClipboard } });
		expect(await copyText("hello")).toBe(true);
		expect(copyToClipboard).toHaveBeenCalledWith("hello");
		expect(clipboardWrite).not.toHaveBeenCalled();
	});

	it("falls back to the browser clipboard, and reports failure", async () => {
		setToolbox({ utils: { copyToClipboard: vi.fn().mockRejectedValue(new Error("x")) } });
		expect(await copyText("a")).toBe(true);
		expect(clipboardWrite).toHaveBeenCalledWith("a");
		clipboardWrite.mockRejectedValue(new Error("denied"));
		setToolbox(undefined);
		expect(await copyText("b")).toBe(false);
	});
});

describe("holdClose", () => {
	it("prevents closing once for overlapping holds and releases after the last", async () => {
		const preventClose = vi.fn().mockResolvedValue(true);
		const releasePreventClose = vi.fn().mockResolvedValue(true);
		setToolbox({ preventClose, releasePreventClose });
		const first = holdClose();
		const second = holdClose();
		expect(preventClose).toHaveBeenCalledTimes(1);
		first();
		first(); // twice is harmless
		expect(releasePreventClose).not.toHaveBeenCalled();
		second();
		expect(releasePreventClose).toHaveBeenCalledTimes(1);
	});

	it("releases when the operation fails, and is a no-op without the API", async () => {
		const preventClose = vi.fn().mockResolvedValue(true);
		const releasePreventClose = vi.fn().mockResolvedValue(true);
		setToolbox({ preventClose, releasePreventClose });
		await expect(whileHoldingClose(async () => Promise.reject(new Error("boom")))).rejects.toThrow("boom");
		expect(releasePreventClose).toHaveBeenCalledTimes(1);
		setToolbox({});
		expect(() => holdClose()()).not.toThrow();
	});
});

describe("canExecuteBatch", () => {
	it("detects dataverseAPI.executeBatch", () => {
		expect(canExecuteBatch()).toBe(false);
		(window as unknown as { dataverseAPI: unknown }).dataverseAPI = { executeBatch: vi.fn() };
		expect(canExecuteBatch()).toBe(true);
	});
});

describe("connectionGuard (#66)", () => {
	it("says no once the tool's connection changes", async () => {
		let current = connection("a", "https://a.crm.dynamics.com/");
		setToolbox({ connections: { getConnection: async () => current } });
		const stillOnConnection = await connectionGuard();
		expect(await stillOnConnection()).toBe(true);
		current = { ...connection("a", "https://a.crm.dynamics.com"), name: "renamed" };
		expect(await stillOnConnection()).toBe(true);
		current = connection("b", "https://b.crm.dynamics.com");
		expect(await stillOnConnection()).toBe(false);
	});
});
