import { describe, expect, it } from "vitest";
import {
	builderWidthLimits,
	clampBuilderWidth,
	clampTreeHeight,
	layoutToPersist,
	normalizePersistedLayout,
	propertiesHost,
	type BuilderPaneState,
} from "./layoutRules";

describe("builderWidthLimits", () => {
	it("uses the narrow default and the 320 floor at 900", () => {
		expect(builderWidthLimits(900)).toEqual({ min: 240, max: 480, defaultValue: 300 });
	});
	it("switches to the full default at exactly 1000", () => {
		expect(builderWidthLimits(999).defaultValue).toBe(300);
		expect(builderWidthLimits(1000)).toEqual({ min: 240, max: 580, defaultValue: 380 });
	});
	it("grows the max with the window at 1600", () => {
		expect(builderWidthLimits(1600)).toEqual({ min: 240, max: 1180, defaultValue: 380 });
	});
	it("never lets the max fall under 320", () => {
		expect(builderWidthLimits(500).max).toBe(320);
	});
});

describe("clamping", () => {
	it("clamps the builder width to the window limits", () => {
		expect(clampBuilderWidth(100, 1600)).toBe(240);
		expect(clampBuilderWidth(400, 1600)).toBe(400);
		expect(clampBuilderWidth(2000, 1600)).toBe(1180);
		expect(clampBuilderWidth(900, 1000)).toBe(580);
	});
	it("clamps the tree height to 120-460", () => {
		expect(clampTreeHeight(0)).toBe(120);
		expect(clampTreeHeight(300)).toBe(300);
		expect(clampTreeHeight(999)).toBe(460);
	});
});

describe("propertiesHost", () => {
	const panes: BuilderPaneState[] = ["docked", "collapsed", "peek"];
	it("is peek for collapsed and peek, whatever the sizes", () => {
		for (const pane of ["collapsed", "peek"] as const) {
			for (const windowWidth of [800, 1000, 1600]) {
				for (const builderWidth of [240, 320, 500]) {
					expect(propertiesHost({ pane, windowWidth, builderWidth })).toBe("peek");
				}
			}
		}
	});
	it("is drawer for a docked pane in a window under 1000", () => {
		expect(propertiesHost({ pane: "docked", windowWidth: 999, builderWidth: 500 })).toBe(
			"drawer",
		);
	});
	it("is docked at exactly a 1000 px window", () => {
		expect(propertiesHost({ pane: "docked", windowWidth: 1000, builderWidth: 500 })).toBe(
			"docked",
		);
	});
	it("is drawer for a docked pane under 320 and docked at exactly 320", () => {
		expect(propertiesHost({ pane: "docked", windowWidth: 1600, builderWidth: 319 })).toBe(
			"drawer",
		);
		expect(propertiesHost({ pane: "docked", windowWidth: 1600, builderWidth: 320 })).toBe(
			"docked",
		);
	});
	it("covers every pane and size combination", () => {
		for (const pane of panes) {
			for (const windowWidth of [999, 1000]) {
				for (const builderWidth of [319, 320]) {
					const expected =
						pane !== "docked"
							? "peek"
							: windowWidth < 1000 || builderWidth < 320
								? "drawer"
								: "docked";
					expect(propertiesHost({ pane, windowWidth, builderWidth })).toBe(expected);
				}
			}
		}
	});
});

describe("normalizePersistedLayout", () => {
	it("keeps valid sizes", () => {
		expect(normalizePersistedLayout({ builderWidth: 400, treeHeight: 200 })).toEqual({
			builderWidth: 400,
			treeHeight: 200,
		});
	});
	it("keeps a wide saved width raw (clamped on use, not here)", () => {
		expect(normalizePersistedLayout({ builderWidth: 1100 })).toEqual({ builderWidth: 1100 });
		expect(clampBuilderWidth(1100, 1000)).toBe(580);
	});
	it("returns an empty object for non-objects", () => {
		for (const raw of [undefined, null, 5, "x", true, []]) {
			expect(normalizePersistedLayout(raw)).toEqual({});
		}
	});
	it("drops non-numbers, NaN, Infinity and out-of-range values", () => {
		expect(normalizePersistedLayout({ builderWidth: "400", treeHeight: Number.NaN })).toEqual(
			{},
		);
		expect(normalizePersistedLayout({ builderWidth: Infinity, treeHeight: -Infinity })).toEqual(
			{},
		);
		expect(normalizePersistedLayout({ builderWidth: 10, treeHeight: 50 })).toEqual({});
		expect(normalizePersistedLayout({ builderWidth: -1, treeHeight: 5000 })).toEqual({});
		expect(normalizePersistedLayout({ builderWidth: 1e9 })).toEqual({});
	});
	it("handles partial objects", () => {
		expect(normalizePersistedLayout({ treeHeight: 300 })).toEqual({ treeHeight: 300 });
		expect(normalizePersistedLayout({ builderWidth: 300, treeHeight: "x" })).toEqual({
			builderWidth: 300,
		});
	});
	it("ignores legacy pane values", () => {
		for (const pane of ["collapsed", "pinned", "peek", "docked", 1, null]) {
			expect(normalizePersistedLayout({ pane, builderWidth: 380 })).toEqual({
				builderWidth: 380,
			});
		}
		expect(normalizePersistedLayout({ pane: "collapsed" })).toEqual({});
	});
});

describe("layoutToPersist", () => {
	it("returns sizes only, never the pane", () => {
		const state = { builderWidth: 350, treeHeight: 200, pane: "collapsed" };
		const out = layoutToPersist(state);
		expect(out).toEqual({ builderWidth: 350, treeHeight: 200 });
		expect(Object.keys(out)).not.toContain("pane");
	});
});
