import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { executeFetchXml, getFetchXmlEntityName } from "./pptbClient";

// #61: FetchXML runs through queryData with the metadata EntitySetName, not the host pluralizer.

type Api = {
	queryData: ReturnType<typeof vi.fn>;
	fetchXmlQuery: ReturnType<typeof vi.fn>;
	getEntityMetadata: ReturnType<typeof vi.fn>;
};

let api: Api;

beforeEach(() => {
	vi.spyOn(console, "log").mockImplementation(() => {});
	vi.spyOn(console, "warn").mockImplementation(() => {});
	api = {
		queryData: vi.fn().mockResolvedValue({ value: [{ name: "x" }] }),
		fetchXmlQuery: vi.fn().mockResolvedValue({ value: [] }),
		getEntityMetadata: vi.fn(async (logicalName: string) => {
			// Metadata from the sandbox org (2026-10-13) for entities the host pluralizer gets wrong:
			// it requests "feedbacks", "msdyn_conversationdatas" and "archivecleanupinfos".
			const sets: Record<string, { EntitySetName: string; LogicalCollectionName: string }> = {
				feedback: { EntitySetName: "feedback", LogicalCollectionName: "feedback" },
				msdyn_conversationdata: { EntitySetName: "msdyn_conversationsdata", LogicalCollectionName: "msdyn_conversationsdata" },
				archivecleanupinfo: { EntitySetName: "ArchiveCleanupInfos", LogicalCollectionName: "archivecleanupinfos" },
			};
			if (!(logicalName in sets)) throw new Error("not found");
			return { LogicalName: logicalName, ...sets[logicalName] };
		}),
	};
	(window as unknown as { dataverseAPI: Api }).dataverseAPI = api;
});

afterEach(() => {
	vi.restoreAllMocks();
	delete (window as unknown as { dataverseAPI?: Api }).dataverseAPI;
});

describe("getFetchXmlEntityName", () => {
	it("reads the root entity, not a link-entity", () => {
		expect(
			getFetchXmlEntityName(
				`<fetch><entity name="feedback"><link-entity name="account" from="accountid" to="regardingobjectid" /></entity></fetch>`,
			),
		).toBe("feedback");
	});

	it("returns undefined for malformed or non-fetch XML", () => {
		expect(getFetchXmlEntityName(`<fetch><entity name="account">`)).toBeUndefined();
		expect(getFetchXmlEntityName(`<savedquery><entity name="account" /></savedquery>`)).toBeUndefined();
		expect(getFetchXmlEntityName(`<fetch />`)).toBeUndefined();
	});
});

describe("executeFetchXml (#61)", () => {
	it("uses queryData with the metadata EntitySetName", async () => {
		const xml = `<fetch top="1"><entity name="feedback"><attribute name="title" /></entity></fetch>`;
		const result = await executeFetchXml(xml);
		expect(api.queryData).toHaveBeenCalledWith(`feedback?fetchXml=${encodeURIComponent(xml)}`);
		expect(api.fetchXmlQuery).not.toHaveBeenCalled();
		expect(result.records).toEqual([{ name: "x" }]);
	});

	it("caches the entity set name between executions", async () => {
		const xml = `<fetch><entity name="msdyn_conversationdata" /></fetch>`;
		await executeFetchXml(xml);
		await executeFetchXml(xml);
		expect(api.getEntityMetadata).toHaveBeenCalledTimes(1);
		expect(api.queryData).toHaveBeenLastCalledWith(`msdyn_conversationsdata?fetchXml=${encodeURIComponent(xml)}`);
	});

	it("falls back to fetchXmlQuery only when metadata is unavailable", async () => {
		const xml = `<fetch><entity name="unknownentity" /></fetch>`;
		await executeFetchXml(xml);
		expect(api.fetchXmlQuery).toHaveBeenCalledWith(xml);
		expect(api.queryData).not.toHaveBeenCalled();
	});

	it("retries with LogicalCollectionName when the EntitySetName segment doesn't route", async () => {
		api.queryData.mockRejectedValueOnce(new Error("0x80060888: Resource not found for the segment 'ArchiveCleanupInfos'."));
		const xml = `<fetch><entity name="archivecleanupinfo" /></fetch>`;
		await executeFetchXml(xml);
		expect(api.queryData).toHaveBeenNthCalledWith(1, `ArchiveCleanupInfos?fetchXml=${encodeURIComponent(xml)}`);
		expect(api.queryData).toHaveBeenNthCalledWith(2, `archivecleanupinfos?fetchXml=${encodeURIComponent(xml)}`);
	});

	it("doesn't fall back when the query itself fails", async () => {
		api.queryData.mockRejectedValueOnce(new Error("0x80041103: attribute not found"));
		await expect(executeFetchXml(`<fetch><entity name="feedback" /></fetch>`)).rejects.toThrow("0x80041103");
		expect(api.fetchXmlQuery).not.toHaveBeenCalled();
	});
});
