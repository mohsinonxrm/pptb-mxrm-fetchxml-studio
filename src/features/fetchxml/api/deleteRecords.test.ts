import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { deleteRecord, deleteRecordsBatch } from "./pptbClient";
import { metadataCache } from "../state/cache";

// #61: record delete uses the metadata EntitySetName, not the host's pluralizer. Sandbox metadata
// (2026-10-12): the host pluralizes adminsettingsentity to "adminsettingsentities"; the real entity
// set is new_admin_settings_entities.

const ID1 = "11111111-1111-1111-1111-111111111111";
const ID2 = "22222222-2222-2222-2222-222222222222";

type Api = Record<string, ReturnType<typeof vi.fn>>;
let api: Api;

beforeEach(() => {
	vi.spyOn(console, "log").mockImplementation(() => {});
	vi.spyOn(console, "warn").mockImplementation(() => {});
	vi.spyOn(console, "error").mockImplementation(() => {});
	api = {
		getEntityMetadata: vi.fn(async (logicalName: string) => ({
			LogicalName: logicalName,
			EntitySetName: "new_admin_settings_entities",
		})),
		executeBatch: vi.fn(async (requests: unknown[]) => requests.map(() => ({ status: 204, headers: {} }))),
		delete: vi.fn().mockResolvedValue(undefined),
	};
	(window as unknown as { dataverseAPI: Api }).dataverseAPI = api;
	(window as unknown as { toolboxAPI: unknown }).toolboxAPI = {
		connections: { getConnection: async () => ({ id: "c1", url: "https://org.crm.dynamics.com", name: "Org" }) },
	};
});

afterEach(() => {
	metadataCache.clear();
	vi.restoreAllMocks();
	delete (window as unknown as { dataverseAPI?: unknown }).dataverseAPI;
	delete (window as unknown as { toolboxAPI?: unknown }).toolboxAPI;
});

describe("record delete (#61)", () => {
	it("deletes a record by its metadata entity set through executeBatch", async () => {
		await deleteRecord("adminsettingsentity", ID1);
		expect(api.executeBatch).toHaveBeenCalledWith([{ method: "DELETE", url: `new_admin_settings_entities(${ID1})` }]);
		expect(api.delete).not.toHaveBeenCalled();
	});

	it("throws the Dataverse error of a failed single delete", async () => {
		api.executeBatch.mockResolvedValue([{ status: 403, headers: {}, body: { error: { code: "0x80040220", message: "No privilege" } } }]);
		await expect(deleteRecord("adminsettingsentity", ID1)).rejects.toThrow("No privilege");
	});

	it("deletes several records in one $batch", async () => {
		const result = await deleteRecordsBatch("adminsettingsentity", [ID1, ID2]);
		expect(result).toEqual({ succeeded: 2, failed: 0, errors: [] });
		expect(api.executeBatch).toHaveBeenCalledTimes(1);
		expect(api.executeBatch.mock.calls[0][0]).toEqual([
			{ method: "DELETE", url: `new_admin_settings_entities(${ID1})` },
			{ method: "DELETE", url: `new_admin_settings_entities(${ID2})` },
		]);
	});

	it("uses the host's delete on hosts without executeBatch (< 1.2.7)", async () => {
		delete api.executeBatch;
		const result = await deleteRecordsBatch("adminsettingsentity", [ID1, ID2]);
		expect(result.succeeded).toBe(2);
		expect(api.delete).toHaveBeenCalledWith("adminsettingsentity", ID1);
		await deleteRecord("adminsettingsentity", ID1);
		expect(api.delete).toHaveBeenCalledTimes(3);
	});

	it("uses the host's delete when metadata is unavailable", async () => {
		api.getEntityMetadata.mockRejectedValue(new Error("not found"));
		await deleteRecord("adminsettingsentity", ID1);
		expect(api.delete).toHaveBeenCalledWith("adminsettingsentity", ID1);
		expect(api.executeBatch).not.toHaveBeenCalled();
	});
});
