import { afterEach, describe, expect, it, vi } from "vitest";
import { resolveLaunchPrefill, viewPrefillNotice } from "./invocation";
import { getViewById, type SavedView } from "./pptbClient";

const id = "0a1b2c3d-4e5f-6a7b-8c9d-0e1f2a3b4c5d";
const view: SavedView = { id, name: "Active Contacts", fetchxml: "<fetch><entity name='contact'/></fetch>", type: "system" };

describe("a view another tool opens FetchXML Studio with (#197)", () => {
	it("loads the whole view, so its layout, execution by ID and Save view overwrite work", async () => {
		const readView = vi.fn(async () => ({ view, entityLogicalName: "contact" }));
		const action = await resolveLaunchPrefill({ viewRef: { id: `{${id.toUpperCase()}}`, entityLogicalName: "savedquery" } }, readView);
		expect(action).toEqual({ kind: "view", view, entityLogicalName: "contact" });
		expect(readView).toHaveBeenCalledWith("system", id.toUpperCase());
	});

	it("reads personal views from userquery", async () => {
		const readView = vi.fn(async () => null);
		await resolveLaunchPrefill({ viewRef: { id, entityLogicalName: "userquery" } }, readView);
		expect(readView).toHaveBeenCalledWith("personal", id);
	});

	it("says so when the view doesn't exist or can't be read, instead of failing silently", async () => {
		const action = await resolveLaunchPrefill({ viewRef: { id, entityLogicalName: "userquery" } }, async () => null);
		expect(action).toEqual({ kind: "viewUnavailable", viewType: "personal", id });
		expect(viewPrefillNotice(action as Extract<typeof action, { kind: "viewUnavailable" }>)).toEqual({
			title: "Couldn't open the view",
			body: `The tool that launched FetchXML Studio asked for a personal view (${id}) that doesn't exist or that you can't read.`,
			type: "error",
		});
	});

	it("reports Dataverse's error when the request fails", async () => {
		const error = { code: "0x80040220", message: "Principal user is missing prvReadSavedQuery privilege" };
		const action = await resolveLaunchPrefill({ viewRef: { id, entityLogicalName: "savedquery" } }, async () => {
			throw error;
		});
		expect(action).toEqual({ kind: "viewUnavailable", viewType: "system", id, error });
		const notice = viewPrefillNotice(action as Extract<typeof action, { kind: "viewUnavailable" }>);
		expect(notice.body).toMatch(/^The tool that launched FetchXML Studio asked for a system view, but Dataverse returned an error: /);
		expect(notice.body).toContain("prvReadSavedQuery");
	});

	it("doesn't query an id that isn't a GUID (#176), and shortens a long one in the notice", async () => {
		const readView = vi.fn();
		const junk = `${id} or savedqueryid ne null ${"x".repeat(60)}`;
		const action = await resolveLaunchPrefill({ viewRef: { id: junk, entityLogicalName: "savedquery" } }, readView);
		expect(readView).not.toHaveBeenCalled();
		expect(action).toMatchObject({ kind: "viewUnavailable", viewType: "system", id: junk });
		const body = viewPrefillNotice(action as Extract<typeof action, { kind: "viewUnavailable" }>).body;
		expect(body).toContain(`(${junk.slice(0, 57)}…)`);
	});

	it("still prefers FetchXML, then a view, then an entity", async () => {
		const readView = vi.fn(async () => ({ view, entityLogicalName: "contact" }));
		expect(await resolveLaunchPrefill({ fetchXml: "<fetch/>", viewRef: { id, entityLogicalName: "savedquery" } }, readView)).toEqual({
			kind: "fetchxml",
			fetchXml: "<fetch/>",
		});
		expect(await resolveLaunchPrefill({ viewRef: { id, entityLogicalName: "savedquery" }, entityLogicalName: "account" }, readView)).toMatchObject({
			kind: "view",
		});
		expect(await resolveLaunchPrefill({ entityLogicalName: " account " }, readView)).toEqual({ kind: "entity", entityLogicalName: "account" });
		expect(await resolveLaunchPrefill(null, readView)).toBeNull();
	});
});

describe("getViewById (#197)", () => {
	afterEach(() => {
		delete (window as { dataverseAPI?: unknown }).dataverseAPI;
	});

	it("reads Load view's columns plus the view's table, by GUID", async () => {
		const queryData = vi.fn(async () => ({
			value: [{ savedqueryid: id, name: "Active Contacts", fetchxml: "<fetch/>", layoutxml: "<grid/>", querytype: 0, iscustomizable: { Value: true }, returnedtypecode: "contact" }],
		}));
		(window as { dataverseAPI?: unknown }).dataverseAPI = { queryData };
		const found = await getViewById("system", id);
		expect(queryData).toHaveBeenCalledWith(
			`savedqueries?$select=savedqueryid,name,description,fetchxml,layoutxml,isdefault,querytype,iscustomizable,ismanaged,solutionid,returnedtypecode&$filter=savedqueryid eq ${id} and componentstate eq 0 and statecode eq 0`,
		);
		expect(found).toMatchObject({ entityLogicalName: "contact", view: { id, name: "Active Contacts", type: "system", layoutxml: "<grid/>", isCustomizable: true } });
	});

	it("is null when nothing comes back or the view has no FetchXML", async () => {
		(window as { dataverseAPI?: unknown }).dataverseAPI = { queryData: async () => ({ value: [] }) };
		expect(await getViewById("personal", id)).toBeNull();
		(window as { dataverseAPI?: unknown }).dataverseAPI = { queryData: async () => ({ value: [{ userqueryid: id, name: "x", fetchxml: " ", returnedtypecode: "contact" }] }) };
		expect(await getViewById("personal", id)).toBeNull();
	});
});
