import { describe, expect, it } from "vitest";
import { buildRowIdentity } from "./rowIdentity";

const A = "11111111-1111-1111-1111-111111111111";
const B = "22222222-2222-2222-2222-222222222222";

describe("buildRowIdentity", () => {
	it("uses the primary key as the row ID for normal results", () => {
		const rows = [{ accountid: A, name: "a" }, { accountid: B, name: "b" }];
		const id = buildRowIdentity(rows, "accountid");
		expect(rows.map(id.getRowId)).toEqual([A, B]);
		expect(id.toRecordIds([A, B])).toEqual([A, B]);
	});

	it("keeps row IDs unique when a 1:N join repeats the primary key", () => {
		const rows = [
			{ accountid: A, "c.fullname": "x" },
			{ accountid: A, "c.fullname": "y" },
			{ accountid: B, "c.fullname": "z" },
		];
		const id = buildRowIdentity(rows, "accountid");
		const rowIds = rows.map(id.getRowId);
		expect(new Set(rowIds).size).toBe(3);
		expect(rowIds[0]).toBe(A);
		expect(rowIds[2]).toBe(B);
		// Selecting both duplicate rows yields the record once.
		expect(id.toRecordIds([rowIds[0], rowIds[1]])).toEqual([A]);
	});

	it("uses position-based IDs when results have no primary key (aggregate/distinct)", () => {
		const rows = [{ n: 3, p: 1 }, { n: 5, p: 2 }];
		const id = buildRowIdentity(rows, "activityid");
		expect(rows.map(id.getRowId)).toEqual(["__row_0", "__row_1"]);
		expect(id.toRecordIds(["__row_0", "__row_1"])).toEqual([]);
	});

	it("is stable across appended pages", () => {
		const page1 = [{ accountid: A }];
		const page2 = [...page1, { accountid: B }];
		expect(buildRowIdentity(page1, "accountid").getRowId(page1[0])).toBe(
			buildRowIdentity(page2, "accountid").getRowId(page2[0]),
		);
	});

	it("never returns synthetic IDs as record IDs", () => {
		const rows = [{ name: "no pk" }, { accountid: A }];
		const id = buildRowIdentity(rows, "accountid");
		expect(id.toRecordIds(rows.map(id.getRowId))).toEqual([A]);
	});
});
