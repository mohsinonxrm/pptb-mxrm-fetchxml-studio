/**
 * Wildcard helper (#98): pattern building, literal escapes (wildcard-characters) and reading a
 * stored value back into the helper's choices.
 */
import { describe, expect, it } from "vitest";
import {
	buildWildcardValue,
	convertWildcardValue,
	escapeWildcards,
	isWildcardOperator,
	parseWildcardValue,
	unescapeWildcards,
	type WildcardValue,
} from "./wildcards";

describe("escapeWildcards / unescapeWildcards", () => {
	it.each([
		["50%", "50[%]"],
		["a_b", "a[_]b"],
		["[x]", "[[]x]"],
		["100% [new]_", "100[%] [[]new][_]"],
		["plain", "plain"],
		["", ""],
	])("%s ↔ %s", (text, escaped) => {
		expect(escapeWildcards(text)).toBe(escaped);
		expect(unescapeWildcards(escaped)).toBe(text);
	});

	it("escapes each character once, including text that looks like an escape", () => {
		expect(escapeWildcards("[%]")).toBe("[[][%]]");
		expect(unescapeWildcards("[[][%]]")).toBe("[%]");
	});
});

describe("buildWildcardValue", () => {
	const v = (match: WildcardValue["match"], text: string, literal = false): WildcardValue => ({ match, text, literal });
	it.each([
		["like", v("contains", "contoso"), "%contoso%"],
		["like", v("begins", "contoso"), "contoso%"],
		["like", v("ends", "contoso"), "%contoso"],
		["like", v("custom", "con_oso%"), "con_oso%"],
		["not-like", v("contains", "50%", true), "%50[%]%"],
		["like", v("begins", "a_b", true), "a[_]b%"],
		["like", v("ends", "[x]", true), "%[[]x]"],
		["like", v("contains", "50%"), "%50%%"],
		// Custom pattern is written as typed: Match literally doesn't apply
		["like", v("custom", "50%", true), "50%"],
		// begins-with / ends-with add their own wildcard
		["begins-with", v("custom", "contoso"), "contoso"],
		["ends-with", v("custom", "50%", true), "50[%]"],
		["not-begin-with", v("contains", "a_b", true), "a[_]b"],
		["not-end-with", v("custom", "a_b"), "a_b"],
	] as const)("%s %j → %s", (operator, value, expected) => {
		expect(buildWildcardValue(operator, value)).toBe(expected);
	});
});

describe("parseWildcardValue", () => {
	it.each([
		["like", "%contoso%", { match: "contains", text: "contoso", literal: false }],
		["like", "contoso%", { match: "begins", text: "contoso", literal: false }],
		["like", "%contoso", { match: "ends", text: "contoso", literal: false }],
		["like", "%50[%]%", { match: "contains", text: "50%", literal: true }],
		["like", "a[_]b%", { match: "begins", text: "a_b", literal: true }],
		["like", "%[[]x]", { match: "ends", text: "[x]", literal: true }],
		["like", "%[%]", { match: "ends", text: "%", literal: true }],
		["not-like", "%a%", { match: "contains", text: "a", literal: false }],
		// Other wildcards open as Custom pattern
		["like", "con_oso%", { match: "custom", text: "con_oso%", literal: false }],
		["like", "%a%b%", { match: "custom", text: "%a%b%", literal: false }],
		["like", "[a-c]%", { match: "custom", text: "[a-c]%", literal: false }],
		["like", "contoso", { match: "custom", text: "contoso", literal: false }],
		["like", "abc[%]", { match: "custom", text: "abc[%]", literal: false }],
		["like", "%", { match: "custom", text: "%", literal: false }],
		["like", "%%", { match: "custom", text: "%%", literal: false }],
		// A new like condition starts as Contains
		["like", undefined, { match: "contains", text: "", literal: false }],
		["like", "", { match: "contains", text: "", literal: false }],
		// begins-with / ends-with: the text, with escapes read back
		["begins-with", "contoso", { match: "custom", text: "contoso", literal: false }],
		["ends-with", "50[%]", { match: "custom", text: "50%", literal: true }],
		["begins-with", "50%", { match: "custom", text: "50%", literal: false }],
		["begins-with", "[a-c]", { match: "custom", text: "[a-c]", literal: false }],
		["begins-with", 123, { match: "custom", text: "123", literal: false }],
	] as const)("%s %j", (operator, value, expected) => {
		expect(parseWildcardValue(operator, value)).toEqual(expected);
	});

	it("round-trips every helper value", () => {
		for (const operator of ["like", "not-like", "begins-with", "ends-with"]) {
			for (const match of ["contains", "begins", "ends"] as const) {
				for (const text of ["contoso", "50% off", "a_b", "[x]", "it's"]) {
					const value = buildWildcardValue(operator, { match, text, literal: true });
					const parsed = parseWildcardValue(operator, value);
					expect(buildWildcardValue(operator, parsed), `${operator} ${value}`).toBe(value);
				}
			}
		}
	});
});

describe("convertWildcardValue", () => {
	it.each([
		["like", "begins-with", "%abc%", "abc"],
		["like", "ends-with", "abc%", "abc"],
		["like", "begins-with", "%50[%]%", "50[%]"],
		["begins-with", "like", "abc", "abc%"],
		["ends-with", "like", "abc", "%abc"],
		["not-end-with", "not-like", "a[_]b", "%a[_]b"],
		["begins-with", "ends-with", "abc", "abc"],
		["like", "not-like", "%abc%", "%abc%"],
		// A custom pattern is kept as written
		["like", "begins-with", "a_b%c", "a_b%c"],
		// Other operators keep the value
		["eq", "like", "abc", "abc"],
		["like", "eq", "%abc%", "%abc%"],
		["like", "begins-with", undefined, undefined],
	])("%s → %s: %s → %s", (from, to, value, expected) => {
		expect(convertWildcardValue(from, to, value)).toBe(expected);
	});
});

describe("isWildcardOperator", () => {
	it("covers the six pattern operators only", () => {
		expect(["like", "not-like", "begins-with", "not-begin-with", "ends-with", "not-end-with"].every(isWildcardOperator)).toBe(true);
		expect(["eq", "ne", "in", "under", "contain-values"].some(isWildcardOperator)).toBe(false);
	});
});
