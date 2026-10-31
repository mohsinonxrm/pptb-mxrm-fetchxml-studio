import { describe, expect, it } from "vitest";
import { viewRefFetchXmlQuery } from "./invocation";

describe("view reference from another tool (#176)", () => {
	const id = "0a1b2c3d-4e5f-6a7b-8c9d-0e1f2a3b4c5d";

	it("reads a system or personal view by its GUID", () => {
		expect(viewRefFetchXmlQuery({ id, entityLogicalName: "savedquery" })).toBe(
			`savedqueries?$select=savedqueryid,fetchxml&$filter=savedqueryid eq ${id}`,
		);
		expect(viewRefFetchXmlQuery({ id: `{${id.toUpperCase()}}`, entityLogicalName: "userquery" })).toBe(
			`userqueries?$select=userqueryid,fetchxml&$filter=userqueryid eq ${id.toUpperCase()}`,
		);
	});

	it("refuses an id that isn't a GUID, so it can't change the $filter", () => {
		expect(viewRefFetchXmlQuery({ id: `${id} or savedqueryid ne null`, entityLogicalName: "savedquery" })).toBeNull();
		expect(viewRefFetchXmlQuery({ id: "", entityLogicalName: "savedquery" })).toBeNull();
	});
});
