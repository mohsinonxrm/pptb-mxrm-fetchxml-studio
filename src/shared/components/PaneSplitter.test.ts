import { act, createElement, useState } from "react";
import { afterEach, beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import { cleanup, render } from "../testing/ui";
import { PaneSplitter } from "./PaneSplitter";

interface Spies {
	onChange: Mock<(value: number) => void>;
	onCommit: Mock<(value: number) => void>;
}

/** Holds the size like a real parent, so successive keys build on each other */
function harness(orientation: "vertical" | "horizontal", spies: Spies, initial = 380) {
	return function Harness() {
		const [value, setValue] = useState(initial);
		return createElement(PaneSplitter, {
			orientation,
			value,
			min: 240,
			max: 600,
			defaultValue: 380,
			onChange: (v: number) => {
				spies.onChange(v);
				setValue(v);
			},
			onCommit: spies.onCommit,
			label: "Resize the builder pane",
			description: "Drag or use arrow keys; double-click resets.",
			controls: "builder-pane",
		});
	};
}

const splitter = () => document.querySelector<HTMLElement>('[role="separator"]')!;

async function press(key: string, shiftKey = false) {
	await act(async () => {
		splitter().dispatchEvent(new KeyboardEvent("keydown", { key, shiftKey, bubbles: true, cancelable: true }));
	});
}

async function pointer(type: string, init: { clientX?: number; clientY?: number; button?: number } = {}) {
	await act(async () => {
		const event = new MouseEvent(type, { bubbles: true, cancelable: true, button: 0, ...init });
		Object.defineProperty(event, "pointerId", { value: 1 });
		splitter().dispatchEvent(event);
	});
}

let spies: Spies;

beforeEach(() => {
	spies = { onChange: vi.fn<(value: number) => void>(), onCommit: vi.fn<(value: number) => void>() };
	Element.prototype.setPointerCapture = vi.fn();
	Element.prototype.releasePointerCapture = vi.fn();
});

afterEach(() => {
	cleanup();
	vi.useRealTimers();
});

describe("PaneSplitter", () => {
	it("exposes the window splitter aria attributes", async () => {
		await render(createElement(harness("vertical", spies)));
		const el = splitter();
		expect(el.getAttribute("aria-orientation")).toBe("vertical");
		expect(el.getAttribute("aria-valuenow")).toBe("380");
		expect(el.getAttribute("aria-valuemin")).toBe("240");
		expect(el.getAttribute("aria-valuemax")).toBe("600");
		expect(el.getAttribute("aria-valuetext")).toBe("380 pixels");
		expect(el.getAttribute("aria-controls")).toBe("builder-pane");
		expect(el.getAttribute("aria-label")).toBe("Resize the builder pane");
		expect(el.tabIndex).toBe(0);
	});

	it("ArrowRight adds 16 and Shift+ArrowRight adds 64", async () => {
		await render(createElement(harness("vertical", spies)));
		await press("ArrowRight");
		expect(spies.onChange).toHaveBeenLastCalledWith(396);
		await press("ArrowRight", true);
		expect(spies.onChange).toHaveBeenLastCalledWith(460);
		await press("ArrowLeft");
		expect(spies.onChange).toHaveBeenLastCalledWith(444);
		expect(splitter().getAttribute("aria-valuenow")).toBe("444");
	});

	it("clamps at min and max", async () => {
		await render(createElement(harness("vertical", spies, 590)));
		await press("ArrowRight", true);
		expect(spies.onChange).toHaveBeenLastCalledWith(600);
		await cleanup();
		await render(createElement(harness("vertical", spies, 250)));
		await press("ArrowLeft", true);
		expect(spies.onChange).toHaveBeenLastCalledWith(240);
	});

	it("Home and End jump to the limits", async () => {
		await render(createElement(harness("vertical", spies)));
		await press("End");
		expect(spies.onChange).toHaveBeenLastCalledWith(600);
		await press("Home");
		expect(spies.onChange).toHaveBeenLastCalledWith(240);
	});

	it("Enter resets to the default", async () => {
		await render(createElement(harness("vertical", spies, 500)));
		await press("Enter");
		expect(spies.onChange).toHaveBeenLastCalledWith(380);
	});

	it("a horizontal splitter uses ArrowUp and ArrowDown, and ignores Left and Right", async () => {
		await render(createElement(harness("horizontal", spies)));
		expect(splitter().getAttribute("aria-orientation")).toBe("horizontal");
		await press("ArrowDown");
		expect(spies.onChange).toHaveBeenLastCalledWith(396);
		await press("ArrowUp", true);
		expect(spies.onChange).toHaveBeenLastCalledWith(332);
		spies.onChange.mockClear();
		await press("ArrowRight");
		await press("ArrowLeft");
		expect(spies.onChange).not.toHaveBeenCalled();
	});

	it("double-click resets to the default and commits", async () => {
		await render(createElement(harness("vertical", spies, 500)));
		await pointer("dblclick");
		expect(spies.onChange).toHaveBeenLastCalledWith(380);
		expect(spies.onCommit).toHaveBeenCalledTimes(1);
		expect(spies.onCommit).toHaveBeenCalledWith(380);
	});

	it("a drag past max clamps and commits exactly once", async () => {
		await render(createElement(harness("vertical", spies)));
		await pointer("pointerdown", { clientX: 100 });
		expect(Element.prototype.setPointerCapture).toHaveBeenCalled();
		expect(document.body.style.userSelect).toBe("none");
		await pointer("pointermove", { clientX: 160 });
		expect(spies.onChange).toHaveBeenLastCalledWith(440);
		await pointer("pointermove", { clientX: 900 });
		expect(spies.onChange).toHaveBeenLastCalledWith(600);
		await pointer("pointermove", { clientX: -900 });
		expect(spies.onChange).toHaveBeenLastCalledWith(240);
		await pointer("pointermove", { clientX: 900 });
		await pointer("pointerup", { clientX: 900 });
		await pointer("lostpointercapture");
		expect(spies.onCommit).toHaveBeenCalledTimes(1);
		expect(spies.onCommit).toHaveBeenCalledWith(600);
		expect(document.body.style.userSelect).not.toBe("none");
	});

	it("lost pointer capture alone commits once", async () => {
		await render(createElement(harness("horizontal", spies)));
		await pointer("pointerdown", { clientY: 50 });
		await pointer("pointermove", { clientY: 70 });
		await pointer("lostpointercapture");
		expect(spies.onCommit).toHaveBeenCalledTimes(1);
		expect(spies.onCommit).toHaveBeenCalledWith(400);
	});

	it("key commits are debounced by 250 ms after the last key", async () => {
		vi.useFakeTimers();
		await render(createElement(harness("vertical", spies)));
		await press("ArrowRight");
		await act(async () => void vi.advanceTimersByTime(200));
		await press("ArrowRight");
		await act(async () => void vi.advanceTimersByTime(200));
		expect(spies.onCommit).not.toHaveBeenCalled();
		await act(async () => void vi.advanceTimersByTime(60));
		expect(spies.onCommit).toHaveBeenCalledTimes(1);
		expect(spies.onCommit).toHaveBeenCalledWith(412);
	});
});
