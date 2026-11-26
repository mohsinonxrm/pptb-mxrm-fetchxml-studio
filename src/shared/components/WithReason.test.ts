/**
 * withReason (#279): the reason describes the control and shows on hover; without one, nothing
 * shows and nothing describes it, though the wrapper stays (so a focused control isn't remounted).
 * A hover while there was nothing to say isn't remembered (#279 review).
 */
import { act, createElement } from "react";
import { Button } from "@fluentui/react-components";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, describedBy, render } from "../testing/ui";
import { withReason } from "./WithReason";

afterEach(() => cleanup());

const wait = (ms: number) => act(() => new Promise<void>((resolve) => setTimeout(resolve, ms)));
/** React makes onPointerEnter / Leave from pointerover / pointerout */
const pointer = async (type: "pointerover" | "pointerout", element: Element | null) => {
	await act(async () => {
		element?.dispatchEvent(
			new PointerEvent(type, { bubbles: true, relatedTarget: document.body }),
		);
	});
	// Past the Tooltip's show and hide delays (250 ms)
	await wait(400);
};
const shown = () => {
	const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
	return !!tip && getComputedStyle(tip).display !== "none";
};
const show = (reason: string | undefined) =>
	render(withReason(createElement(Button, { disabledFocusable: !!reason }, "Save"), reason));

describe("withReason", () => {
	it("a reason describes the control and shows on hover", async () => {
		await show("Enter a name");
		const button = document.querySelector("button");
		expect(describedBy(button)).toBe("Enter a name");
		expect(shown()).toBe(false);
		await pointer("pointerover", button);
		expect(shown()).toBe(true);
		await pointer("pointerout", button);
		expect(shown()).toBe(false);
	});

	it("without a reason, no tooltip shows and nothing describes the control", async () => {
		await show(undefined);
		const button = document.querySelector("button");
		expect(button?.getAttribute("aria-describedby")).toBeNull();
		await pointer("pointerover", button);
		expect(shown()).toBe(false);
	});

	it("a hover while there was nothing to say doesn't pop the reason up later", async () => {
		await show(undefined);
		const button = document.querySelector("button");
		await pointer("pointerover", button);
		await pointer("pointerout", button);
		// The control gains a reason with nothing pointing at it
		await show("Type contact to confirm");
		expect(document.querySelector("button")).toBe(button);
		expect(shown()).toBe(false);
	});
});

describe("withReason: a reason that goes while hovered (#279 review)", () => {
	it("doesn't come back on its own after the pointer left", async () => {
		await show("Enter a name");
		const button = document.querySelector("button");
		await pointer("pointerover", button);
		expect(shown()).toBe(true);
		// The reason goes while the pointer is still on it, then the pointer leaves
		await show(undefined);
		await pointer("pointerout", button);
		await show("Enter a name");
		expect(shown()).toBe(false);
	});
});

describe("withReason: a reason that goes before its tooltip shows (#393 review)", () => {
	it("doesn't latch the tooltip on and show it when the reason comes back", async () => {
		await show("Wait for the query to finish");
		const button = document.querySelector("button");
		await act(async () => {
			button?.dispatchEvent(
				new PointerEvent("pointerover", { bubbles: true, relatedTarget: document.body }),
			);
		});
		// The reason goes within the show delay (250 ms)
		await wait(100);
		await show(undefined);
		await wait(300);
		await pointer("pointerout", button);
		await show("Wait for the query to finish");
		expect(shown()).toBe(false);
	});
});
