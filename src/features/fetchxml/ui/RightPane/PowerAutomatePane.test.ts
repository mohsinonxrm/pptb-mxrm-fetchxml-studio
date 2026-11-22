/**
 * Power Automate on the Code tab (#233): the "List rows" heading once, each field labelled as the
 * connector labels it (a real label for its input), with its hint and its own Copy.
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../../../shared/testing/ui";
import { PowerAutomatePane } from "./PowerAutomatePane";

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

const spec = {
	actionName: "List rows",
	connector: "Microsoft Dataverse connector",
	banner: "In the List rows action, open Show advanced options and paste these values.",
	fields: [
		{ label: "Table name", value: "accounts", hint: "The entity set name (plural)." },
		{ label: "Fetch Xml Query", value: "<fetch>\n</fetch>", multiline: true },
		{ label: "Row count", value: "" },
	],
	notes: ["Top Count comes only from top."],
};

describe("the Power Automate pane", () => {
	it("names the action once, with its connector", async () => {
		await render(createElement(PowerAutomatePane, { spec }));
		const headings = [...document.querySelectorAll("h3")].map((h) => h.textContent);
		expect(headings).toEqual(["List rows"]);
		expect(document.body.textContent).toContain(
			"Microsoft Dataverse connector · labels match the connector",
		);
		expect(document.body.textContent).toContain(spec.banner);
		expect(document.body.textContent).toContain("Top Count comes only from top.");
	});

	it("labels each value as the connector does, with its hint", async () => {
		await render(createElement(PowerAutomatePane, { spec }));
		// Field labels a native input with <label for>
		const labelOf = (control: HTMLElement | null) =>
			control?.id ? document.querySelector(`label[for="${control.id}"]`)?.textContent : null;
		const tableName = document.querySelector<HTMLInputElement>("input");
		expect(labelOf(tableName)).toBe("Table name");
		expect(tableName?.value).toBe("accounts");
		expect(tableName?.readOnly).toBe(true);
		const describedBy = tableName?.getAttribute("aria-describedby");
		expect(describedBy && document.getElementById(describedBy)?.textContent).toBe(
			"The entity set name (plural).",
		);
		const fetchXml = document.querySelector<HTMLTextAreaElement>("textarea");
		expect(labelOf(fetchXml)).toBe("Fetch Xml Query");
	});

	it("has a Copy per field, off with the reason when the field is empty", async () => {
		await render(createElement(PowerAutomatePane, { spec }));
		const copy = (name: string) =>
			document.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`);
		expect(copy("Copy Table name")?.getAttribute("aria-disabled")).toBeNull();
		expect(copy("Copy Fetch Xml Query")).toBeTruthy();
		expect(copy("Copy Row count: it's empty")?.getAttribute("aria-disabled")).toBe("true");
		// Copy all is in the Code tab's row, not here
		expect(
			[...document.querySelectorAll("button")].some((b) =>
				b.textContent?.includes("Copy all"),
			),
		).toBe(false);
	});
});
