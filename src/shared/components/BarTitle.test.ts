/**
 * A MessageBar's title with a visually hidden stop (#186 PR 3): the bar's text reads "Title. Body",
 * and a title with its own punctuation gets a space only.
 */
import { createElement } from "react";
import { MessageBar, MessageBarBody } from "@fluentui/react-components";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render } from "../testing/ui";
import { BarTitle } from "./BarTitle";

afterEach(() => cleanup());

const barText = async (title: string, body: string) => {
	await render(
		createElement(
			MessageBar,
			null,
			createElement(MessageBarBody, null, createElement(BarTitle, null, title), body),
		),
	);
	return document.querySelector(".fui-MessageBarBody")?.textContent;
};

describe("BarTitle", () => {
	it("ends the title with a full stop the bar reads, but doesn't show", async () => {
		expect(await barText("Export failed", "Disk full")).toBe("Export failed. Disk full");
		const stop = document.querySelector(".fui-MessageBarTitle span");
		expect(stop?.textContent).toBe(". ");
		expect(stop?.className).toBeTruthy();
	});

	it("a title with its own punctuation gets a space only", async () => {
		expect(await barText("Still loading…", "Wait")).toBe("Still loading… Wait");
		cleanup();
		expect(await barText("Heads up:", "Slow")).toBe("Heads up: Slow");
	});
});
