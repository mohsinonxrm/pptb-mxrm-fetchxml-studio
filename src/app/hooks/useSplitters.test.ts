/**
 * useSplitters (#301): the defaults, the tree/properties drag (applied only between 30 and 80 % of
 * the left pane), the builder-width drag (applied only between 300 and 800 px), mouseup ending a
 * drag, and the document listeners removed on unmount.
 */
import { act, createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { useSplitters, type Splitters } from "./useSplitters";

let latest: Splitters;

/** The same wiring as AppShell: the app root, the left pane and the two handles */
function Probe() {
	latest = useSplitters();
	return createElement(
		"div",
		{ "data-app-root": true },
		createElement(
			"div",
			{ "data-left-pane": true },
			createElement("div", {
				"data-testid": "horizontal",
				onMouseDown: latest.startHorizontalDrag,
			}),
		),
		createElement("div", { "data-testid": "vertical", onMouseDown: latest.startVerticalDrag }),
	);
}

/** jsdom lays nothing out, so give the element the rect the drag measures against */
function setRect(selector: string, rect: { left?: number; top?: number; height?: number }) {
	const element = document.querySelector(selector) as HTMLElement;
	element.getBoundingClientRect = () =>
		({ left: 0, top: 0, width: 0, height: 0, ...rect }) as DOMRect;
}

async function mouse(
	target: EventTarget,
	type: "mousedown" | "mousemove" | "mouseup",
	init: MouseEventInit = {},
) {
	await act(async () => {
		target.dispatchEvent(new MouseEvent(type, { bubbles: true, ...init }));
	});
}

const handle = (name: "horizontal" | "vertical") =>
	document.querySelector(`[data-testid="${name}"]`) as HTMLElement;

beforeEach(async () => {
	await render(createElement(Probe));
	// Left pane: top at 100 px, 500 px high. App root: left edge at 20 px.
	setRect("[data-left-pane]", { top: 100, height: 500 });
	setRect("[data-app-root]", { left: 20 });
});

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("useSplitters", () => {
	it("starts at 58 % tree height and a 480 px builder, with no drag", () => {
		expect(latest.topHeight).toBe(58);
		expect(latest.leftPaneWidth).toBe(480);
		expect(latest.isDraggingHorizontal).toBe(false);
		expect(latest.isDraggingVertical).toBe(false);
	});

	it("ignores mousemove until a handle is pressed", async () => {
		await mouse(document, "mousemove", { clientX: 620, clientY: 350 });
		expect(latest.topHeight).toBe(58);
		expect(latest.leftPaneWidth).toBe(480);
	});

	it("a horizontal drag sets the tree height, applied only between 30 and 80 %", async () => {
		await mouse(handle("horizontal"), "mousedown");
		expect(latest.isDraggingHorizontal).toBe(true);
		expect(latest.isDraggingVertical).toBe(false);

		await mouse(document, "mousemove", { clientY: 350 }); // 250 / 500
		expect(latest.topHeight).toBe(50);

		await mouse(document, "mousemove", { clientY: 550 }); // 90 %: outside, not applied
		expect(latest.topHeight).toBe(50);
		await mouse(document, "mousemove", { clientY: 200 }); // 20 %: outside, not applied
		expect(latest.topHeight).toBe(50);

		await mouse(document, "mousemove", { clientY: 250 }); // 30 %: the lower limit applies
		expect(latest.topHeight).toBe(30);
		await mouse(document, "mousemove", { clientY: 500 }); // 80 %: the upper limit applies
		expect(latest.topHeight).toBe(80);

		// The builder width doesn't move
		expect(latest.leftPaneWidth).toBe(480);
	});

	it("a vertical drag sets the builder width, applied only between 300 and 800 px", async () => {
		await mouse(handle("vertical"), "mousedown");
		expect(latest.isDraggingVertical).toBe(true);
		expect(latest.isDraggingHorizontal).toBe(false);

		await mouse(document, "mousemove", { clientX: 620 }); // 620 - 20
		expect(latest.leftPaneWidth).toBe(600);

		await mouse(document, "mousemove", { clientX: 920 }); // 900 px: outside, not applied
		expect(latest.leftPaneWidth).toBe(600);
		await mouse(document, "mousemove", { clientX: 270 }); // 250 px: outside, not applied
		expect(latest.leftPaneWidth).toBe(600);

		await mouse(document, "mousemove", { clientX: 320 }); // 300 px: the lower limit applies
		expect(latest.leftPaneWidth).toBe(300);
		await mouse(document, "mousemove", { clientX: 820 }); // 800 px: the upper limit applies
		expect(latest.leftPaneWidth).toBe(800);

		// The tree height doesn't move
		expect(latest.topHeight).toBe(58);
	});

	it("mouseup ends the drag; later mousemoves change nothing", async () => {
		await mouse(handle("horizontal"), "mousedown");
		await mouse(document, "mousemove", { clientY: 350 });
		await mouse(document, "mouseup");
		expect(latest.isDraggingHorizontal).toBe(false);
		await mouse(document, "mousemove", { clientY: 450 });
		expect(latest.topHeight).toBe(50);

		await mouse(handle("vertical"), "mousedown");
		await mouse(document, "mousemove", { clientX: 620 });
		await mouse(document, "mouseup");
		expect(latest.isDraggingVertical).toBe(false);
		await mouse(document, "mousemove", { clientX: 720 });
		expect(latest.leftPaneWidth).toBe(600);
	});

	it("removes its document listeners on unmount", async () => {
		const add = vi.spyOn(document, "addEventListener");
		const remove = vi.spyOn(document, "removeEventListener");

		await mouse(handle("horizontal"), "mousedown");
		await mouse(handle("vertical"), "mousedown");
		const added = add.mock.calls.filter(([type]) => type === "mousemove" || type === "mouseup");
		expect(added).toHaveLength(4);

		cleanup();

		for (const [type, listener] of added) {
			expect(remove).toHaveBeenCalledWith(type, listener);
		}
	});
});
