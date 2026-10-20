import { describe, expect, it, vi } from "vitest";
import { DATAVERSE_DEFAULT_PAGE_SIZE, fetchPageWithRecovery, halvePageSize } from "./pageRecovery";
import type { PagingState } from "./paging";

// #68: a page over the PPTB host's 16 MB limit is fetched again with half the count.

const TOO_LARGE = new Error("Dataverse response exceeds the maximum size");

const paging = (overrides: Partial<PagingState> = {}): PagingState => ({
	mode: "fetchxml",
	xml: "<fetch><entity name='account'><all-attributes /></entity></fetch>",
	currentPage: 1,
	nextPage: 2,
	moreRecords: true,
	isRetrieveAllInProgress: true,
	retrieveAll: true,
	...overrides,
});

describe("halvePageSize (#68)", () => {
	it("halves Dataverse's default page size and keeps the page with a paging cookie", () => {
		const next = halvePageSize(
			paging({ fullPageSize: 5000, pagingCookie: '<cookie page="1" />' }),
		);
		expect(next?.pageSize).toBe(2500);
		expect(next?.nextPage).toBe(2);
	});

	it("halves a first page that set no count from Dataverse's default, keeping page 1 (#183)", () => {
		const next = halvePageSize(paging({ pageSize: DATAVERSE_DEFAULT_PAGE_SIZE, currentPage: 0, nextPage: 1 }));
		expect(next?.pageSize).toBe(2500);
		expect(next?.nextPage).toBe(1);
	});

	it("renumbers the page for simple paging so it starts at the same row", () => {
		// Page 3 of 5,000 starts at row 10,001: at 2,500 per page that's page 5
		const next = halvePageSize(paging({ pageSize: 5000, currentPage: 2, nextPage: 3 }));
		expect(next?.pageSize).toBe(2500);
		expect(next?.nextPage).toBe(5);
		expect((next!.nextPage - 1) * next!.pageSize!).toBe((3 - 1) * 5000);
	});

	it("doesn't shrink saved views, unknown or odd sizes, or below the minimum", () => {
		expect(halvePageSize(paging({ mode: "view", fullPageSize: 5000 }))).toBeUndefined();
		expect(halvePageSize(paging())).toBeUndefined();
		expect(halvePageSize(paging({ pageSize: 625 }))).toBeUndefined();
		expect(halvePageSize(paging({ pageSize: 150 }))).toBeUndefined();
	});
});

describe("fetchPageWithRecovery (#68)", () => {
	it("retries the page with half the count until it fits, and returns that position", async () => {
		const fetchPage = vi
			.fn()
			.mockRejectedValueOnce(TOO_LARGE)
			.mockRejectedValueOnce(TOO_LARGE)
			.mockResolvedValueOnce({ records: [] });
		const { result, state } = await fetchPageWithRecovery(
			paging({ fullPageSize: 5000, pagingCookie: "c" }),
			fetchPage,
		);
		expect(result).toEqual({ records: [] });
		expect(state.pageSize).toBe(1250);
		expect(fetchPage.mock.calls.map(([s]) => (s as PagingState).pageSize)).toEqual([
			undefined,
			2500,
			1250,
		]);
	});

	it("stops after the maximum halvings and throws the host's error", async () => {
		const fetchPage = vi.fn().mockRejectedValue(TOO_LARGE);
		await expect(
			fetchPageWithRecovery(paging({ pageSize: 4000, pagingCookie: "c" }), fetchPage, {
				maxHalvings: 2,
			}),
		).rejects.toBe(TOO_LARGE);
		expect(fetchPage).toHaveBeenCalledTimes(3);
	});

	it("doesn't shrink for other errors", async () => {
		const error = new Error("0x80040220: missing prvReadAccount privilege");
		const fetchPage = vi.fn().mockRejectedValue(error);
		await expect(fetchPageWithRecovery(paging({ pageSize: 5000 }), fetchPage)).rejects.toBe(
			error,
		);
		expect(fetchPage).toHaveBeenCalledTimes(1);
	});

	it("waits and retries after service protection errors (#96)", async () => {
		const sleep = vi.fn().mockResolvedValue(undefined);
		const fetchPage = vi
			.fn()
			.mockRejectedValueOnce(
				new Error("0x80072326: Number of concurrent requests exceeded the limit of 52."),
			)
			.mockResolvedValueOnce("page");
		const { result } = await fetchPageWithRecovery(paging(), fetchPage, { retry: { sleep } });
		expect(result).toBe("page");
		expect(sleep).toHaveBeenCalledTimes(1);
	});
});
