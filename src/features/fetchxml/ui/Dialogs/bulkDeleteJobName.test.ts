import { describe, expect, it } from "vitest";
import { defaultBulkDeleteJobName } from "./bulkDeleteJobName";

describe("defaultBulkDeleteJobName (#182)", () => {
	it("names the entity and the time given", () => {
		const at = new Date(2026, 10, 2, 9, 30, 0);
		expect(defaultBulkDeleteJobName("Contact", at)).toBe(`Bulk Delete Contact - ${at.toLocaleString()}`);
	});

	it("says Records when the entity has no display name", () => {
		const at = new Date(2026, 10, 2, 9, 30, 0);
		expect(defaultBulkDeleteJobName(undefined, at)).toBe(`Bulk Delete Records - ${at.toLocaleString()}`);
		expect(defaultBulkDeleteJobName("", at)).toBe(`Bulk Delete Records - ${at.toLocaleString()}`);
	});
});
