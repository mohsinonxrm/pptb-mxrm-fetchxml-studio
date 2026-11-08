import { describe, expect, it } from "vitest";
import type { Advisory, AdvisoryLevel } from "../../model/advisories";
import { advisoriesTabLabel, resultsTabLabel } from "./tabLabels";

const make = (...levels: AdvisoryLevel[]): Advisory[] =>
	levels.map((level, i) => ({ id: `a${i}`, level, title: "t", detail: "d" }) as Advisory);

describe("resultsTabLabel", () => {
	it("has no badge before the first run", () => {
		expect(resultsTabLabel(undefined, "en-US")).toEqual({
			ariaLabel: "Results",
			menuLabel: "Results",
		});
	});
	it("shows 0 after a run", () => {
		expect(resultsTabLabel(0, "en-US")).toEqual({
			badge: "0",
			ariaLabel: "Results, 0 rows",
			menuLabel: "Results (0 rows)",
		});
	});
	it("uses the singular for one row", () => {
		expect(resultsTabLabel(1, "en-US")).toEqual({
			badge: "1",
			ariaLabel: "Results, 1 row",
			menuLabel: "Results (1 row)",
		});
	});
	it("shows the exact count, formatted, with no cap", () => {
		expect(resultsTabLabel(48213, "en-US")).toEqual({
			badge: "48,213",
			ariaLabel: "Results, 48,213 rows",
			menuLabel: "Results (48,213 rows)",
		});
	});
	it("formats with the given locale", () => {
		expect(resultsTabLabel(48213, "de-DE").badge).toBe("48.213");
	});
});

describe("advisoriesTabLabel", () => {
	it("has no badge when empty", () => {
		const label = advisoriesTabLabel([]);
		expect(label).toEqual({ ariaLabel: "Advisories", menuLabel: "Advisories" });
		expect(label.badgeColor).toBeUndefined();
	});
	it("describes a single level", () => {
		expect(advisoriesTabLabel(make("advice"))).toEqual({
			badge: "1",
			badgeColor: "informative",
			ariaLabel: "Advisories, 1: 1 performance advice",
			menuLabel: "Advisories (1)",
		});
	});
	it("breaks mixed levels down, worst first, skipping zero levels", () => {
		expect(advisoriesTabLabel(make("advice", "warning", "error", "warning"))).toEqual({
			badge: "4",
			badgeColor: "danger",
			ariaLabel:
				"Advisories, 4: 1 will be rejected, 2 might be rejected, 1 performance advice",
			menuLabel: "Advisories (4)",
		});
		expect(advisoriesTabLabel(make("error", "advice")).ariaLabel).toBe(
			"Advisories, 2: 1 will be rejected, 1 performance advice",
		);
	});
	it("colours by the worst level present", () => {
		expect(advisoriesTabLabel(make("error", "warning", "advice")).badgeColor).toBe("danger");
		expect(advisoriesTabLabel(make("warning", "advice")).badgeColor).toBe("warning");
		expect(advisoriesTabLabel(make("warning")).badgeColor).toBe("warning");
		expect(advisoriesTabLabel(make("advice", "advice")).badgeColor).toBe("informative");
	});
	it("keeps the level labels the same for many", () => {
		expect(advisoriesTabLabel(make("error", "error", "error")).ariaLabel).toBe(
			"Advisories, 3: 3 will be rejected",
		);
	});
});
