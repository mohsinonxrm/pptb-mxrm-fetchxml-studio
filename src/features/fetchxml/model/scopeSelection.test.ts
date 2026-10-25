import { describe, expect, it } from "vitest";
import {
	environmentKey,
	readRememberedScopes,
	rememberedScopeFor,
	restorableIds,
	withRememberedScope,
} from "./scopeSelection";

describe("remembered publisher / solution selection (#25)", () => {
	it("keys environments by URL, ignoring case and trailing slashes", () => {
		expect(environmentKey("https://Contoso.crm.dynamics.com/")).toBe("https://contoso.crm.dynamics.com");
		expect(environmentKey(" https://contoso.crm.dynamics.com// ")).toBe("https://contoso.crm.dynamics.com");
		expect(environmentKey(undefined)).toBe("");
	});

	it("keeps each environment's selection separately", () => {
		let stored: unknown = undefined;
		stored = withRememberedScope(stored, "https://a", { publisherIds: ["p1"], solutionIds: ["s1"] });
		stored = withRememberedScope(stored, "https://b", { publisherIds: ["p2"], solutionIds: [] });
		expect(rememberedScopeFor(stored, "https://a")).toEqual({ publisherIds: ["p1"], solutionIds: ["s1"] });
		expect(rememberedScopeFor(stored, "https://b")).toEqual({ publisherIds: ["p2"], solutionIds: [] });
		expect(rememberedScopeFor(stored, "https://c")).toBeUndefined();
		expect(rememberedScopeFor(stored, "")).toBeUndefined();
	});

	it("forgets an environment whose selection is cleared, and never stores without a key", () => {
		const stored = withRememberedScope(undefined, "https://a", { publisherIds: ["p1"], solutionIds: [] });
		expect(withRememberedScope(stored, "https://a", { publisherIds: [], solutionIds: [] })).toEqual({});
		expect(withRememberedScope(undefined, "", { publisherIds: ["p1"], solutionIds: [] })).toEqual({});
	});

	it("ignores a stored value that isn't a selection map", () => {
		expect(readRememberedScopes(null)).toEqual({});
		expect(readRememberedScopes("x")).toEqual({});
		expect(readRememberedScopes([1, 2])).toEqual({});
		expect(readRememberedScopes({ "https://a": { publisherIds: ["p1", 3, ""], solutionIds: "s1" }, "https://b": null })).toEqual({
			"https://a": { publisherIds: ["p1"], solutionIds: [] },
		});
	});

	it("restores only the ids that still exist, in the remembered order", () => {
		expect(restorableIds(["s2", "gone", "s1"], ["s1", "s2", "s3"])).toEqual(["s2", "s1"]);
		expect(restorableIds(["gone"], [])).toEqual([]);
	});
});
