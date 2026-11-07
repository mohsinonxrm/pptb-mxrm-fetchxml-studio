import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { clearOtherDefaultPublicViews } from "./pptbClient";

// #307: only one active public view can be the default, but Dataverse doesn't clear the previous
// default when another view is created or updated with isdefault (verified on a sandbox,
// 2026-11-02), so FetchXML Studio clears the entity's other default public views itself.

const NEW_VIEW = "f93dd770-2bd7-f111-a05d-70a8a5afaf52";
const OLD_DEFAULT = "00000000-0000-0000-00aa-000010001050";
const SECOND = "c0bb354c-6de2-4873-96ff-237e8b2d6b21";

type Api = Record<string, ReturnType<typeof vi.fn>>;
let api: Api;

beforeEach(() => {
	vi.spyOn(console, "log").mockImplementation(() => {});
	vi.spyOn(console, "error").mockImplementation(() => {});
	api = {
		queryData: vi.fn(async () => ({ value: [{ savedqueryid: OLD_DEFAULT, name: "All Territories" }] })),
		update: vi.fn(async () => undefined),
	};
	(window as unknown as { dataverseAPI: Api }).dataverseAPI = api;
});

afterEach(() => {
	vi.restoreAllMocks();
	delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
});

describe("clearOtherDefaultPublicViews (#307)", () => {
	it("reads the entity's other default public views and clears isdefault on each", async () => {
		expect(await clearOtherDefaultPublicViews("territory", NEW_VIEW)).toEqual([]);
		const query = api.queryData.mock.calls[0][0] as string;
		expect(query).toContain("returnedtypecode eq 'territory'");
		expect(query).toContain("querytype eq 0");
		expect(query).toContain("statecode eq 0");
		expect(query).toContain("isdefault eq true");
		expect(query).toContain(`savedqueryid ne ${NEW_VIEW}`);
		expect(api.update).toHaveBeenCalledTimes(1);
		expect(api.update).toHaveBeenCalledWith("savedquery", OLD_DEFAULT, { isdefault: false });
	});

	it("does nothing more when the new view is the only default", async () => {
		api.queryData.mockResolvedValue({ value: [] });
		expect(await clearOtherDefaultPublicViews("territory", NEW_VIEW)).toEqual([]);
		expect(api.update).not.toHaveBeenCalled();
	});

	it("keeps going after a failure and reports each view it couldn't clear", async () => {
		api.queryData.mockResolvedValue({
			value: [
				{ savedqueryid: OLD_DEFAULT, name: "All Territories" },
				{ savedqueryid: SECOND, name: "Managed default" },
			],
		});
		api.update.mockImplementation(async (_entity: string, id: string) => {
			if (id === SECOND) throw new Error("0x80040220: The component can't be customized");
		});
		const notCleared = await clearOtherDefaultPublicViews("territory", NEW_VIEW);
		expect(api.update).toHaveBeenCalledTimes(2);
		expect(notCleared).toHaveLength(1);
		expect(notCleared[0]).toMatchObject({ id: SECOND, name: "Managed default" });
		expect(notCleared[0].error).toMatch(/can't be customized/);
	});

	it("escapes the entity name, strips braces from the view ID, and refuses a value that isn't an ID", async () => {
		await clearOtherDefaultPublicViews("o'brien", `{${NEW_VIEW}}`);
		const query = api.queryData.mock.calls[0][0] as string;
		expect(query).toContain("returnedtypecode eq 'o''brien'");
		expect(query).toContain(`savedqueryid ne ${NEW_VIEW}`);
		await expect(clearOtherDefaultPublicViews("territory", "not-a-guid) or (1 eq 1")).rejects.toThrow(/Not a view ID/);
	});
});
