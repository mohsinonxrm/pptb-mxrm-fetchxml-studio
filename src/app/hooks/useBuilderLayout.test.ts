/**
 * useBuilderLayout (#234): the default sizes for the window, the saved sizes (read once, clamped on
 * use, never written back clamped), saving only settled sizes and only after the read, and where
 * Properties live (the drawer when the window is under 1000 px or the pane under 320 px).
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, flush, render } from "../../shared/testing/ui";
import { LAYOUT_KEY, useBuilderLayout, type BuilderLayout } from "./useBuilderLayout";

let latest: BuilderLayout;
function Probe({ forceCollapsed }: { forceCollapsed?: boolean } = {}) {
	latest = useBuilderLayout({ forceCollapsed });
	return null;
}

const settings = {
	get: vi.fn<(key: string) => Promise<unknown>>(),
	set: vi.fn<(key: string, value: unknown) => Promise<void>>(),
};
const originalWidth = window.innerWidth;
const setWindowWidth = async (width: number) => {
	Object.defineProperty(window, "innerWidth", { configurable: true, value: width });
	await act(async () => {
		window.dispatchEvent(new Event("resize"));
	});
};

beforeEach(async () => {
	settings.get.mockResolvedValue(undefined);
	settings.set.mockResolvedValue(undefined);
	(window as unknown as { toolboxAPI: unknown }).toolboxAPI = { settings };
	Object.defineProperty(window, "innerWidth", { configurable: true, value: 1280 });
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
	settings.get.mockReset();
	settings.set.mockReset();
	delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
	Object.defineProperty(window, "innerWidth", { configurable: true, value: originalWidth });
});

describe("the builder pane's sizes", () => {
	it("start at the design's defaults: 380 px wide (300 under a 1000 px window), a 240 px tree", async () => {
		await render(createElement(Probe));
		await flush(3);
		expect(latest.builderWidth).toBe(380);
		expect(latest.treeHeight).toBe(240);
		expect(latest.builderWidthLimits).toEqual({ min: 240, max: 860, defaultValue: 380 });
		cleanup();
		Object.defineProperty(window, "innerWidth", { configurable: true, value: 900 });
		await render(createElement(Probe));
		await flush(3);
		expect(latest.builderWidth).toBe(300);
	});

	it("come back from the saved layout, and a saved width is clamped to the window, not rewritten", async () => {
		settings.get.mockResolvedValue({ builderWidth: 700, treeHeight: 300, pane: "collapsed" });
		await render(createElement(Probe));
		await flush(5);
		expect(settings.get).toHaveBeenCalledWith(LAYOUT_KEY);
		expect(latest.builderWidth).toBe(700);
		expect(latest.treeHeight).toBe(300);
		// A smaller window: the width is clamped (max 1000 - 420 = 580), nothing is saved
		await setWindowWidth(1000);
		expect(latest.builderWidth).toBe(580);
		expect(settings.set).not.toHaveBeenCalled();
	});

	it("ignore a saved layout that isn't one", async () => {
		settings.get.mockResolvedValue({ builderWidth: "wide", treeHeight: -5 });
		await render(createElement(Probe));
		await flush(5);
		expect(latest.builderWidth).toBe(380);
		expect(latest.treeHeight).toBe(240);
	});

	it("follow a drag at once, and are saved only when it settles", async () => {
		await render(createElement(Probe));
		await flush(5);
		await act(async () => latest.setBuilderWidth(420));
		expect(latest.builderWidth).toBe(420);
		expect(settings.set).not.toHaveBeenCalled();
		await act(async () => latest.commitBuilderWidth(440));
		expect(settings.set).toHaveBeenCalledTimes(1);
		expect(settings.set).toHaveBeenLastCalledWith(LAYOUT_KEY, { builderWidth: 440 });
		await act(async () => latest.commitTreeHeight(200));
		expect(settings.set).toHaveBeenLastCalledWith(LAYOUT_KEY, {
			builderWidth: 440,
			treeHeight: 200,
		});
	});

	it("aren't saved before the saved layout has been read, and what was set meanwhile wins", async () => {
		let resolve: (value: unknown) => void = () => {};
		settings.get.mockReturnValue(new Promise((r) => (resolve = r)));
		await render(createElement(Probe));
		await act(async () => latest.commitBuilderWidth(500));
		expect(settings.set).not.toHaveBeenCalled();
		await act(async () => resolve({ builderWidth: 300, treeHeight: 300 }));
		await flush(3);
		expect(latest.builderWidth).toBe(500);
		expect(latest.treeHeight).toBe(300);
		// What was settled meanwhile is saved once the read is done, with the saved tree height
		expect(settings.set).toHaveBeenCalledTimes(1);
		expect(settings.set).toHaveBeenLastCalledWith(LAYOUT_KEY, {
			builderWidth: 500,
			treeHeight: 300,
		});
		await act(async () => latest.commitTreeHeight(260));
		expect(settings.set).toHaveBeenLastCalledWith(LAYOUT_KEY, {
			builderWidth: 500,
			treeHeight: 260,
		});
	});
});

describe("when the saved layout can't be read", () => {
	it("the sizes still work, but nothing is saved over the ones that couldn't be read", async () => {
		vi.spyOn(console, "warn").mockImplementation(() => {});
		settings.get.mockRejectedValue(new Error("settings unavailable"));
		await render(createElement(Probe));
		await flush(5);
		await act(async () => latest.commitBuilderWidth(450));
		expect(latest.builderWidth).toBe(450);
		expect(settings.set).not.toHaveBeenCalled();
	});
});

describe("where Properties live", () => {
	it("below the tree, in the drawer under a 1000 px window, or when the pane is under 320 px", async () => {
		await render(createElement(Probe));
		await flush(3);
		expect(latest.propertiesHost).toBe("docked");
		await setWindowWidth(999);
		expect(latest.propertiesHost).toBe("drawer");
		await setWindowWidth(1280);
		expect(latest.propertiesHost).toBe("docked");
		await act(async () => latest.setBuilderWidth(319));
		expect(latest.propertiesHost).toBe("drawer");
		await act(async () => latest.setBuilderWidth(320));
		expect(latest.propertiesHost).toBe("docked");
	});

	it("the drawer is open only while Properties live in it", async () => {
		await render(createElement(Probe));
		await flush(3);
		await act(async () => latest.setInspectorOpen(true));
		expect(latest.inspectorOpen).toBe(false);
		await setWindowWidth(900);
		await act(async () => latest.setInspectorOpen(true));
		expect(latest.inspectorOpen).toBe(true);
		await setWindowWidth(1280);
		expect(latest.inspectorOpen).toBe(false);
		// Back in drawer mode, it stays closed until asked for again
		await setWindowWidth(900);
		expect(latest.propertiesHost).toBe("drawer");
		expect(latest.inspectorOpen).toBe(false);
	});
});

describe("collapsing the pane (#234)", () => {
	it("starts docked; collapses to the rail, peeks at either part, and docks again", async () => {
		await render(createElement(Probe));
		await flush(3);
		expect(latest.pane).toBe("docked");
		expect(latest.peekPart).toBeNull();
		await act(async () => latest.collapse());
		expect(latest.pane).toBe("collapsed");
		expect(latest.propertiesHost).toBe("peek");
		await act(async () => latest.openPeek("tree"));
		expect(latest.pane).toBe("peek");
		expect(latest.peekPart).toBe("tree");
		await act(async () => latest.openPeek("props"));
		expect(latest.peekPart).toBe("props");
		await act(async () => latest.closePeek());
		expect(latest.pane).toBe("collapsed");
		expect(latest.peekPart).toBeNull();
		await act(async () => latest.openPeek("tree"));
		await act(async () => latest.dock());
		expect(latest.pane).toBe("docked");
		expect(latest.peekPart).toBeNull();
		expect(latest.propertiesHost).toBe("docked");
		// A peek asked for while docked doesn't open
		await act(async () => latest.openPeek("tree"));
		expect(latest.pane).toBe("docked");
		expect(latest.peekPart).toBeNull();
	});

	it("peeks at the pane's width, at least 340 px, within the window", async () => {
		await render(createElement(Probe));
		await flush(3);
		expect(latest.peekWidth).toBe(380);
		await act(async () => latest.setBuilderWidth(600));
		expect(latest.peekWidth).toBe(600);
		await setWindowWidth(900);
		await act(async () => latest.setBuilderWidth(260));
		expect(latest.peekWidth).toBe(340);
		// 400 - 44 - 48 = 308
		await setWindowWidth(400);
		expect(latest.peekWidth).toBe(308);
	});

	it("closes the drawer when the pane collapses", async () => {
		await setWindowWidth(900);
		await render(createElement(Probe));
		await flush(3);
		await act(async () => latest.setInspectorOpen(true));
		expect(latest.inspectorOpen).toBe(true);
		await act(async () => latest.collapse());
		expect(latest.inspectorOpen).toBe(false);
		await act(async () => latest.dock());
		expect(latest.propertiesHost).toBe("drawer");
		expect(latest.inspectorOpen).toBe(false);
	});

	it("stays collapsed while forced; a peek still opens; docking applies once the force ends", async () => {
		await render(createElement(Probe, { forceCollapsed: true }));
		await flush(3);
		expect(latest.pane).toBe("collapsed");
		await act(async () => latest.openPeek("props"));
		expect(latest.pane).toBe("peek");
		await act(async () => latest.dock());
		expect(latest.pane).toBe("collapsed");
		await render(createElement(Probe, { forceCollapsed: false }));
		await flush(2);
		expect(latest.pane).toBe("docked");
		// Collapsed by choice stays collapsed when a force ends
		await act(async () => latest.collapse());
		await render(createElement(Probe, { forceCollapsed: true }));
		await render(createElement(Probe, { forceCollapsed: false }));
		await flush(2);
		expect(latest.pane).toBe("collapsed");
	});
});
