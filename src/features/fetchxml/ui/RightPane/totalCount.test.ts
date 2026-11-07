/** The results info bar's total record count (#98, count-rows). */
import { describe, expect, it } from "vitest";
import { entityTotalText, totalCountText } from "./totalCount";
import { totalFromRecordCountCollection } from "../../api/pptbClient";

describe("totalCountText", () => {
	it.each([
		[1234, false, 50, "1,234 total"],
		[0, false, 0, "0 total"],
		[5000, false, 50, "5,000 total"],
		// limit exceeded: the limit is the lower bound (5,000 standard, 500 elastic)
		[5000, true, 50, "5,000+ total"],
		[500, true, 50, "500+ total"],
		// more rows loaded than the limit (Retrieve all)
		[5000, true, 12000, "12,000+ total"],
	] as const)("%d, exceeded %s, %d rows → %s", (total, exceeded, rows, expected) => {
		expect(totalCountText(total, exceeded, rows)).toBe(expected);
	});

	it("is undefined without a count (not requested: -1, or a view)", () => {
		expect(totalCountText(-1, false, 50)).toBeUndefined();
		expect(totalCountText(undefined, undefined, 50)).toBeUndefined();
	});
});

describe("entityTotalText", () => {
	it("says it's unfiltered and a snapshot", () => {
		expect(entityTotalText("account", 158965)).toBe(
			"About 158,965 account records in all: a snapshot from the last 24 hours that ignores this query's filters (RetrieveTotalRecordCount).",
		);
	});
});

describe("totalFromRecordCountCollection", () => {
	// Shape verified live: RetrieveTotalRecordCount(EntityNames=@p)?@p=["account"]
	const collection = { Count: 2, IsReadOnly: false, Keys: ["account", "contact"], Values: [158965, 127533] };
	it("finds the entity's count", () => {
		expect(totalFromRecordCountCollection(collection, "contact")).toBe(127533);
		expect(totalFromRecordCountCollection(collection, "Account")).toBe(158965);
	});
	it("is undefined when the entity isn't there", () => {
		expect(totalFromRecordCountCollection(collection, "lead")).toBeUndefined();
		expect(totalFromRecordCountCollection(undefined, "account")).toBeUndefined();
		expect(totalFromRecordCountCollection({ Keys: ["account"] }, "account")).toBeUndefined();
	});
});
