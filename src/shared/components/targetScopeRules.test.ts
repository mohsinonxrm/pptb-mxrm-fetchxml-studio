import { describe, expect, it } from "vitest";
import {
	TYPE_TO_CONFIRM_RECORD_THRESHOLD,
	describeScope,
	isProductionEnvironment,
	isTypedConfirmation,
	requiresTypedConfirmation,
} from "./targetScopeRules";

// #69: confirmations name the environment and the scope

describe("describeScope", () => {
	it("counts the selected records exactly", () => {
		expect(describeScope({ kind: "records", count: 1 })).toBe("1 record");
		expect(describeScope({ kind: "records", count: 25 })).toBe("25 records");
	});

	it("never claims a total it doesn't know for every matching record", () => {
		expect(describeScope({ kind: "allMatching" })).toBe("Every record the query matches");
		expect(describeScope({ kind: "allMatching", loadedCount: 50, moreRecords: true })).toBe(
			"Every record the query matches (more than the 50 records loaded)",
		);
		expect(describeScope({ kind: "allMatching", loadedCount: 7, moreRecords: false })).toBe(
			"Every record the query matches (7 records when the query last ran)",
		);
		expect(describeScope({ kind: "allMatching", loadedCount: 1 })).toBe(
			"Every record the query matches (1 record loaded; there may be more)",
		);
	});

	it("passes worded scopes through", () => {
		expect(describeScope({ kind: "text", text: "A new personal view" })).toBe("A new personal view");
	});
});

describe("requiresTypedConfirmation", () => {
	it("asks on Production, whatever the count", () => {
		expect(isProductionEnvironment("Production")).toBe(true);
		expect(requiresTypedConfirmation("Production", { kind: "records", count: 1 })).toBe(true);
	});

	it("asks above the threshold and when the total is unknown", () => {
		expect(requiresTypedConfirmation("Dev", { kind: "records", count: TYPE_TO_CONFIRM_RECORD_THRESHOLD })).toBe(false);
		expect(requiresTypedConfirmation("Dev", { kind: "records", count: TYPE_TO_CONFIRM_RECORD_THRESHOLD + 1 })).toBe(true);
		expect(requiresTypedConfirmation(undefined, { kind: "allMatching", loadedCount: 3 })).toBe(true);
		expect(requiresTypedConfirmation("UAT", { kind: "text", text: "A view" })).toBe(false);
	});

	it("accepts the expected word ignoring case and spaces", () => {
		expect(isTypedConfirmation(" Account ", "account")).toBe(true);
		expect(isTypedConfirmation("accounts", "account")).toBe(false);
		expect(isTypedConfirmation("", "")).toBe(false);
	});
});
