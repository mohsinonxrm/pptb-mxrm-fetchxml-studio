/** Record search for the hierarchy value picker (#98). */
import { describe, expect, it } from "vitest";
import { buildRecordNameQuery, buildRecordSearchQuery, isGuid, normalizeGuid } from "./recordSearch";

const ACCOUNT = { entitySetName: "accounts", idAttribute: "accountid", nameAttribute: "name" };

describe("buildRecordSearchQuery", () => {
	it("lists the first records by name when there's no text", () => {
		expect(buildRecordSearchQuery(ACCOUNT, "  ")).toBe("accounts?$select=accountid,name&$orderby=name&$top=10");
	});

	it("searches the start of the primary name", () => {
		expect(buildRecordSearchQuery(ACCOUNT, "Fourth")).toBe(
			"accounts?$select=accountid,name&$filter=startswith(name,'Fourth')&$orderby=name&$top=10",
		);
	});

	it("doubles quotes, escapes wildcards and URL-encodes the literal", () => {
		const query = buildRecordSearchQuery(ACCOUNT, "O'Brien 50%_[a] & #1");
		const encoded = query.match(/startswith\(name,'(.*)'\)/)![1];
		expect(decodeURIComponent(encoded)).toBe("O''Brien 50[%][_][[]a] & #1");
		// Nothing in the literal can end the query option or the URL
		expect(encoded).not.toMatch(/[ &#+]/);
		expect(query.split("&").map((p) => p.split("=")[0])).toEqual(["accounts?$select", "$filter", "$orderby", "$top"]);
	});

	it("takes another top", () => {
		expect(buildRecordSearchQuery(ACCOUNT, "", 5)).toMatch(/\$top=5$/);
	});
});

describe("buildRecordNameQuery", () => {
	it("reads one record's name, braces removed", () => {
		expect(buildRecordNameQuery(ACCOUNT, "{5D5DB393-BE8E-F111-8076-000D3A3BD26B}")).toBe(
			"accounts(5D5DB393-BE8E-F111-8076-000D3A3BD26B)?$select=name",
		);
	});
});

describe("isGuid / normalizeGuid", () => {
	it.each([
		["5d5db393-be8e-f111-8076-000d3a3bd26b", true],
		["{5D5DB393-BE8E-F111-8076-000D3A3BD26B}", true],
		["  5d5db393-be8e-f111-8076-000d3a3bd26b ", true],
		["{5d5db393-be8e-f111-8076-000d3a3bd26b", false],
		["5d5db393be8ef1118076000d3a3bd26b", false],
		["not-a-guid", false],
		["", false],
	])("%s → %s", (value, expected) => {
		expect(isGuid(value)).toBe(expected);
	});

	it("strips braces and spaces", () => {
		expect(normalizeGuid(" {5d5db393-be8e-f111-8076-000d3a3bd26b} ")).toBe("5d5db393-be8e-f111-8076-000d3a3bd26b");
	});
});
