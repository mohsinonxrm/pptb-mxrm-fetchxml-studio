/**
 * Compare to › Another attribute says what's known about the condition attribute's type (#279,
 * owner: an error only when Dataverse is known not to compare it): while its metadata loads, when
 * it couldn't be read, for a type no pairing is verified for, and for multi-select choices.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, installHost, removeHost, render } from "../../../../../../shared/testing/ui";
import { ValueOfPicker, type ValueOfTypeStatus } from "./ValueOfPicker";

afterEach(() => {
	cleanup();
	removeHost();
});

const show = (typeStatus: ValueOfTypeStatus, typeLabel?: string, valueof = "name") =>
	render(
		createElement(ValueOfPicker, {
			fetchQuery: null,
			rootEntityName: "account",
			effectiveEntityName: "account",
			onRootRecord: true,
			scope: { rootName: "account", aliases: [], laterAliases: [] },
			valueof,
			compatibleTypes: typeStatus === "ok" ? ["String", "Memo"] : [],
			typeStatus,
			typeLabel,
			onChange: () => {},
		}),
	);
const text = () => document.body.textContent ?? "";
/** The attribute picker is the second combobox (the first is the Entity dropdown) */
const attributePicker = () =>
	[...document.querySelectorAll<HTMLElement>('[role="combobox"]')].at(-1);
const pickerOff = () =>
	attributePicker()?.hasAttribute("disabled") ||
	attributePicker()?.getAttribute("aria-disabled") === "true";

describe("ValueOfPicker and the attribute's type", () => {
	it("while the type is read: says so, no error, and waits", async () => {
		installHost();
		await show("loading");
		expect(text()).toContain("Loading the attribute…");
		expect(text()).not.toContain("can't compare");
		expect(pickerOff()).toBe(true);
	});

	it("a type it couldn't read: a warning, and every attribute can be picked", async () => {
		installHost();
		await show("unknown");
		expect(text()).toContain("couldn't read this attribute's type");
		expect(text()).not.toContain("Dataverse can't compare");
		expect(pickerOff()).toBe(false);
	});

	it("a type no pairing is verified for: a warning naming it, not an error", async () => {
		installHost();
		await show("unverified", "Image");
		expect(text()).toContain(
			"Comparing an Image attribute to another attribute isn't verified to work. Dataverse decides when the query runs.",
		);
		expect(pickerOff()).toBe(false);
	});

	it("multi-select choices: Dataverse can't compare them, so the picker is off", async () => {
		installHost();
		await show("incomparable", "MultiSelectPicklist");
		expect(text()).toContain(
			"Dataverse can't compare MultiSelectPicklist attributes with another attribute.",
		);
		expect(pickerOff()).toBe(true);
	});

	it("a verified type: no message once an attribute is chosen", async () => {
		installHost();
		await show("ok");
		expect(text()).not.toContain("isn't verified");
		expect(text()).not.toContain("can't compare");
		expect(pickerOff()).toBe(false);
	});
});
