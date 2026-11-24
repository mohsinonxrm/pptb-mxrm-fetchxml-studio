/**
 * withReason (#279): the reason describes the control and shows on hover; without one, nothing
 * shows and nothing describes it, though the wrapper stays (so a focused control isn't remounted).
 */
import { act, createElement } from "react";
import { Button } from "@fluentui/react-components";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, describedBy, render } from "../testing/ui";
import { withReason } from "./WithReason";

afterEach(() => cleanup());

const hover = async (element: Element | null) => {
	await act(async () => {
		element?.dispatchEvent(new PointerEvent("pointerenter", { bubbles: true }));
		element?.dispatchEvent(new MouseEvent("mouseenter", { bubbles: false }));
		await new Promise((resolve) => setTimeout(resolve, 400));
	});
};
const tooltip = () => document.querySelector('[role="tooltip"]');

describe("withReason", () => {
	it("a reason describes the control and shows on hover", async () => {
		await render(
			withReason(createElement(Button, { disabledFocusable: true }, "Save"), "Enter a name"),
		);
		const button = document.querySelector("button");
		expect(describedBy(button)).toBe("Enter a name");
		await hover(button);
		expect(tooltip()?.textContent).toBe("Enter a name");
	});

	it("without a reason, no tooltip shows and nothing describes the control", async () => {
		await render(withReason(createElement(Button, null, "Save"), undefined));
		const button = document.querySelector("button");
		expect(button?.getAttribute("aria-describedby")).toBeNull();
		await hover(button);
		expect(tooltip()).toBeNull();
	});
});
