/**
 * Inputs that can't be used say why where you'd look (#279): Match literally with a custom
 * pattern, and the time without a date.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, click, render } from "../../../../../../shared/testing/ui";
import { WildcardValueField } from "./WildcardValueField";
import { DateTimeInputPicker } from "../../../../../../shared/components/DateTimeInputPicker";
import type { AttributeMetadata } from "../../../../api/pptbClient";

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("field hints", () => {
	it("Match literally: a custom pattern says why it can't be ticked", async () => {
		await render(
			createElement(WildcardValueField, {
				operator: "like",
				value: "con%",
				onChange: () => {},
			}),
		);
		const custom = [...document.querySelectorAll<HTMLInputElement>('input[type="radio"]')].find(
			(r) => r.closest("span, div")?.textContent?.includes("Custom pattern"),
		);
		await click(custom);
		expect(document.body.textContent).toContain(
			"A custom pattern is sent as written: put a wildcard in [ ] to match it as text.",
		);
	});

	it("the time asks for a date first", async () => {
		vi.spyOn(console, "warn").mockImplementation(() => {});
		await render(
			createElement(DateTimeInputPicker, {
				attribute: {
					LogicalName: "createdon",
					AttributeType: "DateTime",
					AttributeTypeName: { Value: "DateTimeType" },
				} as unknown as AttributeMetadata,
				onChange: () => {},
			}),
		);
		expect(
			[...document.querySelectorAll<HTMLInputElement>("input")].some(
				(i) => i.placeholder === "Pick a date first",
			),
		).toBe(true);
	});
});
