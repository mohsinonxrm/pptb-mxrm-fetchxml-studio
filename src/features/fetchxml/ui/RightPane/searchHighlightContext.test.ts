import { describe, expect, it } from "vitest";
import { searchMatches } from "./searchHighlightContext";

describe("searchMatches", () => {
	it("finds nothing for an empty or blank search", () => {
		expect(searchMatches("Contoso", "")).toEqual([]);
		expect(searchMatches("Contoso", "   ")).toEqual([]);
	});

	it("matches case-insensitively, ignoring surrounding spaces", () => {
		expect(searchMatches("Contoso Ltd", " LTD ")).toEqual([[8, 11]]);
	});

	it("finds every occurrence", () => {
		expect(searchMatches("$300,300.00", "300")).toEqual([
			[1, 4],
			[5, 8],
		]);
	});

	it("doesn't overlap matches", () => {
		expect(searchMatches("aaaa", "aa")).toEqual([
			[0, 2],
			[2, 4],
		]);
		expect(searchMatches("aaa", "aa")).toEqual([[0, 2]]);
	});

	it("finds nothing when the text doesn't contain the search", () => {
		expect(searchMatches("Fabrikam", "contoso")).toEqual([]);
	});
});
