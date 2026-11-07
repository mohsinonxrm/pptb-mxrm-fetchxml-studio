import { act, createElement, useEffect } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { button, cleanup, click, flush, installHost, removeHost, render, toasts } from "../testing/ui";
import { NotifyProvider } from "./NotifyProvider";
import { useCopy } from "./useCopy";
import { useNotify, type NotifyApi } from "./useNotify";

let api: NotifyApi | undefined;

function Probe() {
	const value = useNotify();
	useEffect(() => {
		api = value;
	}, [value]);
	return null;
}

async function mount(resetKey = "a") {
	await render(createElement(NotifyProvider, { resetKey }, createElement(Probe)));
}

async function advance(ms: number) {
	await act(async () => {
		await vi.advanceTimersByTimeAsync(ms);
	});
	await flush();
}

async function show(fn: (notify: NotifyApi["notify"]) => void) {
	await act(async () => {
		fn(api!.notify);
	});
	await flush();
}

beforeEach(() => {
	vi.useFakeTimers();
	api = undefined;
});

afterEach(() => {
	cleanup();
	removeHost();
	vi.useRealTimers();
});

describe("NotifyProvider", () => {
	it("shows the title and body", async () => {
		await mount();
		await show((n) => n.success({ title: 'Saved "Active accounts"', body: "Not published yet" }));
		expect(toasts()).toEqual(['Saved "Active accounts"Not published yet']);
	});

	it("removes a timed toast after its timeout", async () => {
		await mount();
		await show((n) => n.success("Copied 3 rows"));
		expect(toasts()).toHaveLength(1);
		await advance(5000);
		expect(toasts()).toHaveLength(0);
	});

	it("runs the action and dismisses the toast", async () => {
		await mount();
		const onClick = vi.fn();
		await show((n) => n.info({ title: "Deleted 2 records", action: { label: "Undo", onClick } }));
		await click(button("Undo", document));
		expect(onClick).toHaveBeenCalledTimes(1);
		await advance(1000);
		expect(toasts()).toHaveLength(0);
	});

	it("keeps an error toast, with Dismiss, until it's dismissed", async () => {
		await mount();
		await show((n) => n.error({ title: "The query didn't run" }));
		await advance(60000);
		expect(toasts()).toHaveLength(1);
		await click(button("Dismiss", document));
		await advance(1000);
		expect(toasts()).toHaveLength(0);
	});

	it("shows the same message twice as one toast", async () => {
		await mount();
		await show((n) => n.success("Copied 3 rows"));
		await show((n) => n.success("Copied 3 rows"));
		await show((n) => n.success("Copied 3 rows"));
		await advance(1000);
		expect(toasts()).toEqual(["Copied 3 rows"]);
	});

	it("replaces a toast that has the same key", async () => {
		await mount();
		await show((n) => n.info({ title: "Loading 100 rows…", key: "load" }));
		await show((n) => n.info({ title: "Loading 200 rows…", key: "load" }));
		await advance(1000);
		expect(toasts()).toEqual(["Loading 200 rows…"]);
	});

	it("queues a 4th toast until one goes", async () => {
		await mount();
		await show((n) => {
			n.error("One");
			n.error("Two");
			n.error("Three");
			n.error("Four");
		});
		expect(toasts()).toEqual(["OneDismiss", "TwoDismiss", "ThreeDismiss"]);
		await click(button("Dismiss", document));
		await advance(1000);
		expect(toasts().some((t) => t.startsWith("Four"))).toBe(true);
	});

	it("clears every toast when resetKey changes", async () => {
		await mount("a");
		await show((n) => {
			n.error("One");
			n.error("Two");
		});
		expect(toasts()).toHaveLength(2);
		await mount("a");
		expect(toasts()).toHaveLength(2);
		await mount("b");
		await advance(1000);
		expect(toasts()).toHaveLength(0);
	});

	it("dismisses by key", async () => {
		await mount();
		await show((n) => n.error({ title: "One", key: "k" }));
		await act(async () => api!.dismiss("k"));
		await advance(1000);
		expect(toasts()).toHaveLength(0);
	});
});

describe("useNotify outside the provider", () => {
	it("doesn't throw and shows nothing", async () => {
		await render(createElement(Probe));
		expect(api).toBeDefined();
		expect(() => {
			api!.notify.success("x");
			api!.notify("error", { title: "y" }).dismiss();
			api!.dismiss("k");
			api!.dismissAll();
		}).not.toThrow();
		expect(toasts()).toHaveLength(0);
	});
});

describe("useCopy", () => {
	function CopyProbe({ onDone }: { onDone: (ok: boolean) => void }) {
		const copy = useCopy();
		return createElement("button", { onClick: () => void copy("text", "the job URL").then(onDone) }, "Copy");
	}

	it("toasts 'Copied …' and resolves true when the clipboard works", async () => {
		installHost();
		const done = vi.fn();
		await render(createElement(NotifyProvider, null, createElement(CopyProbe, { onDone: done })));
		await click(button("Copy", document));
		expect(toasts()).toEqual(["Copied the job URL"]);
		expect(done).toHaveBeenCalledWith(true);
	});

	it("toasts an error and resolves false when it doesn't", async () => {
		const host = installHost();
		(window as unknown as { toolboxAPI: { utils: { copyToClipboard: () => Promise<void> } } }).toolboxAPI.utils.copyToClipboard = async () => {
			throw new Error("denied");
		};
		vi.spyOn(console, "warn").mockImplementation(() => undefined);
		Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async () => Promise.reject(new Error("no")) } });
		void host;
		const done = vi.fn();
		await render(createElement(NotifyProvider, null, createElement(CopyProbe, { onDone: done })));
		await click(button("Copy", document));
		expect(toasts()[0]).toContain("Couldn't copy to the clipboard");
		expect(done).toHaveBeenCalledWith(false);
	});
});
