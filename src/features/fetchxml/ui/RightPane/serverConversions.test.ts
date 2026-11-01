import { describe, expect, it } from "vitest";
import { conversionKey, needsConversion, serverConversionFor } from "./serverConversions";

describe("Code tab: Dataverse conversions (#199)", () => {
	it("only the C# QueryExpression and SQL formats call Dataverse", () => {
		expect(serverConversionFor("csharp-queryexpr")).toBe("queryexpression");
		expect(serverConversionFor("sql")).toBe("sql");
		for (const format of ["csharp-fetchexpr", "javascript", "pac", "powerautomate", "webapi"]) {
			expect(serverConversionFor(format)).toBeUndefined();
		}
	});

	it("keys a result by environment and query", () => {
		const xml = '<fetch><entity name="account" /></fetch>';
		expect(conversionKey(xml, "https://a.crm.dynamics.com")).not.toBe(conversionKey(xml, "https://b.crm.dynamics.com"));
		expect(conversionKey(xml, "https://a.crm.dynamics.com")).toBe(conversionKey(xml, "https://a.crm.dynamics.com"));
		expect(conversionKey(xml, undefined)).not.toBe(conversionKey(`${xml} `, undefined));
	});

	it("requests a query once: not again while it's pending or after it succeeded", () => {
		const key = conversionKey("<fetch />", "https://a");
		expect(needsConversion(key, undefined, undefined)).toBe(true);
		// Pending: requested, no result yet (or a result for an older query)
		expect(needsConversion(key, key, undefined)).toBe(false);
		expect(needsConversion(key, key, { key: "older", code: "x", error: null })).toBe(false);
		// Done
		expect(needsConversion(key, key, { key, code: "var query = …", error: null })).toBe(false);
	});

	it("requests again for a changed query, or to retry a failed one", () => {
		const key = conversionKey("<fetch />", "https://a");
		const changed = conversionKey("<fetch top='5' />", "https://a");
		expect(needsConversion(changed, key, { key, code: "x", error: null })).toBe(true);
		expect(needsConversion(key, key, { key, code: "", error: "Throttled" })).toBe(true);
	});
});
