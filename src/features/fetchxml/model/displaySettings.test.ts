import { describe, expect, it } from "vitest";
import { defaultDisplaySettings, GRID_DENSITY_LABELS, normalizeGridDensity, treeSizeFor } from "./displaySettings";
import { mergeWithDefaults } from "../../../shared/hooks/usePersistedSettings";

describe("grid and tree density (#228)", () => {
	it("defaults to medium", () => {
		expect(defaultDisplaySettings.gridDensity).toBe("medium");
	});

	it("offers Medium, Small and Extra small, in that order", () => {
		expect(Object.entries(GRID_DENSITY_LABELS)).toEqual([
			["medium", "Medium"],
			["small", "Small"],
			["extra-small", "Extra small"],
		]);
	});

	it("keeps known densities and falls back to medium for anything else", () => {
		expect(normalizeGridDensity("small")).toBe("small");
		expect(normalizeGridDensity("extra-small")).toBe("extra-small");
		for (const v of ["medium", "compact", "", undefined, null, 3, {}]) {
			expect(normalizeGridDensity(v)).toBe("medium");
		}
	});

	it("uses the small tree for Small and Extra small (Fluent's Tree has no extra-small)", () => {
		expect(treeSizeFor("medium")).toBe("medium");
		expect(treeSizeFor("small")).toBe("small");
		expect(treeSizeFor("extra-small")).toBe("small");
	});

	it("loads saved settings from before #228 with medium, and repairs an unknown saved value", () => {
		expect(mergeWithDefaults({ useLogicalNames: true }).gridDensity).toBe("medium");
		expect(mergeWithDefaults({ gridDensity: "huge" }).gridDensity).toBe("medium");
		expect(mergeWithDefaults({ gridDensity: "extra-small" }).gridDensity).toBe("extra-small");
		expect(mergeWithDefaults(null)).toEqual(defaultDisplaySettings);
	});
});
