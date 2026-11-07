import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	CONNECTION_CHANGED_MESSAGE,
	STATE_CONNECTION_CHANGED_MESSAGE,
	batchResultError,
	buildStatePatchRequest,
	updateRecordStatesInBatches,
	deleteRecordsInBatches,
	isRecordNotFound,
	recordUrl,
	type BatchRequest,
	type BatchResult,
} from "./batchDelete";

// #61: record delete through executeBatch with explicit EntitySetName URLs (PPTB ≥ 1.2.7)

const ok = (): BatchResult => ({ status: 204, statusText: "No Content", headers: {} });
const fail = (status: number, code: string, message: string): BatchResult => ({
	status,
	headers: {},
	body: { error: { code, message } },
});
const notFound = (id: string) => fail(404, "0x80040217", `new_admin_settings_entity With Id = ${id} Does Not Exist`);
const ids = (n: number) => Array.from({ length: n }, (_, i) => `00000000-0000-0000-0000-${String(i).padStart(12, "0")}`);

const noWait = () => Promise.resolve();

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	vi.restoreAllMocks();
});

describe("recordUrl", () => {
	it("builds a relative entity set URL and strips braces", () => {
		expect(recordUrl("new_admin_settings_entities", "{ABC}")).toBe("new_admin_settings_entities(ABC)");
	});
});

describe("batchResultError / isRecordNotFound", () => {
	it("reads the Dataverse error from the body", () => {
		expect(batchResultError(fail(403, "0x80040220", "No delete privilege"))).toEqual({ code: "0x80040220", message: "No delete privilege" });
		expect(batchResultError({ status: 500, statusText: "Server Error", headers: {} }).message).toBe("HTTP 500 Server Error");
	});

	it("treats only a missing record as not found, not an unknown entity set", () => {
		expect(isRecordNotFound(notFound("x"))).toBe(true);
		expect(isRecordNotFound(fail(404, "0x80060888", "Resource not found for the segment 'adminsettingsentities'."))).toBe(false);
		expect(isRecordNotFound(fail(400, "0x80040217", "Does Not Exist"))).toBe(false);
	});
});

describe("deleteRecordsInBatches", () => {
	it("sends DELETE requests with the given entity set, in batches", async () => {
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		const progress = vi.fn();
		const result = await deleteRecordsInBatches("new_admin_settings_entities", ids(5), {
			executeBatch,
			batchSize: 2,
			onProgress: progress,
			wait: noWait,
		});
		expect(result).toEqual({ succeeded: 5, failed: 0, errors: [] });
		expect(executeBatch).toHaveBeenCalledTimes(3);
		expect(executeBatch.mock.calls[0][0]).toEqual([
			{ method: "DELETE", url: `new_admin_settings_entities(${ids(5)[0]})` },
			{ method: "DELETE", url: `new_admin_settings_entities(${ids(5)[1]})` },
		]);
		expect(progress).toHaveBeenLastCalledWith(expect.objectContaining({ completed: 5, total: 5, batchesCompleted: 3, totalBatches: 3 }));
	});

	it("reports per-record failures from the results", async () => {
		const [a, b] = ids(2);
		const executeBatch = vi.fn(async () => [ok(), fail(403, "0x80040220", "No delete privilege")]);
		const result = await deleteRecordsInBatches("accounts", [a, b], { executeBatch, wait: noWait });
		expect(result).toEqual({ succeeded: 1, failed: 1, errors: [`${b}: No delete privilege`] });
	});

	it("redoes a batch that stopped early one record at a time; already deleted records count as deleted", async () => {
		const [a, b, c] = ids(3);
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) throw new Error("Malformed Dataverse batch response: missing operation results");
			const url = requests[0].url;
			if (url.includes(a)) return [notFound(a)]; // deleted by the stopped batch
			if (url.includes(b)) return [fail(400, "0x80040265", "Cascade blocked")];
			return [ok()];
		});
		const result = await deleteRecordsInBatches("accounts", [a, b, c], { executeBatch, wait: noWait });
		expect(executeBatch).toHaveBeenCalledTimes(4);
		expect(result).toEqual({ succeeded: 2, failed: 1, errors: [`${b}: Cascade blocked`] });
	});

	it("doesn't count a 404 as deleted on the first attempt", async () => {
		const [a] = ids(1);
		const executeBatch = vi.fn(async () => [notFound(a)]);
		const result = await deleteRecordsInBatches("accounts", [a], { executeBatch, wait: noWait });
		expect(result.failed).toBe(1);
	});

	it("retries a throttled batch after waiting", async () => {
		const wait = vi.fn(noWait);
		const executeBatch = vi
			.fn<(requests: BatchRequest[]) => Promise<BatchResult[]>>()
			.mockRejectedValueOnce(new Error("0x80072322: Number of requests exceeded the limit of 6000 over time window of 300 seconds."))
			.mockImplementation(async (requests) => requests.map(ok));
		const result = await deleteRecordsInBatches("accounts", ids(2), { executeBatch, wait });
		expect(wait).toHaveBeenCalledTimes(1);
		expect(result.succeeded).toBe(2);
	});

	it("doesn't resend a throttled batch after the connection changed during the wait (#175)", async () => {
		const wait = vi.fn(noWait);
		const executeBatch = vi
			.fn<(requests: BatchRequest[]) => Promise<BatchResult[]>>()
			.mockRejectedValueOnce(new Error("0x80072322: Number of requests exceeded the limit of 6000 over time window of 300 seconds."))
			.mockImplementation(async (requests) => requests.map(ok));
		let checks = 0;
		// On the connection before the batch, changed by the time the wait is over
		const result = await deleteRecordsInBatches("accounts", ids(2), { executeBatch, wait, shouldContinue: () => ++checks === 1 });
		expect(wait).toHaveBeenCalledTimes(1);
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(result.succeeded).toBe(0);
		expect(result.failed).toBe(2);
		expect(result.errors.every((e) => e.endsWith(CONNECTION_CHANGED_MESSAGE))).toBe(true);
	});

	it("reports a lone record's error without retrying it", async () => {
		const [a] = ids(1);
		const executeBatch = vi.fn(async () => {
			throw new Error("HTTP 401");
		});
		const result = await deleteRecordsInBatches("accounts", [a], { executeBatch, wait: noWait });
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(result).toEqual({ succeeded: 0, failed: 1, errors: [`${a}: HTTP 401`] });
	});

	it("stops when the connection changes and reports the rest as not deleted (#66)", async () => {
		const all = ids(4);
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		let checks = 0;
		const result = await deleteRecordsInBatches("accounts", all, {
			executeBatch,
			batchSize: 2,
			shouldContinue: () => ++checks === 1,
			wait: noWait,
		});
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(result.succeeded).toBe(2);
		expect(result.failed).toBe(2);
		expect(result.errors).toEqual([`${all[2]}: ${CONNECTION_CHANGED_MESSAGE}`, `${all[3]}: ${CONNECTION_CHANGED_MESSAGE}`]);
	});
});

// #143: Activate / Deactivate through the same runner: PATCH { statecode, statuscode }
describe("updateRecordStatesInBatches", () => {
	const target = { statecode: 1, statuscode: 2 };

	it("builds a PATCH of statecode and statuscode to the record URL", () => {
		expect(buildStatePatchRequest("accounts", "{ABC}", target)).toEqual({
			method: "PATCH",
			url: "accounts(ABC)",
			body: { statecode: 1, statuscode: 2 },
		});
	});

	it("sends PATCH requests in batches and reports progress", async () => {
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		const progress = vi.fn();
		const all = ids(3);
		const result = await updateRecordStatesInBatches("accounts", all, target, { executeBatch, batchSize: 2, onProgress: progress, wait: noWait });
		expect(result).toEqual({ succeeded: 3, failed: 0, errors: [] });
		expect(executeBatch.mock.calls[0][0]).toEqual([
			{ method: "PATCH", url: `accounts(${all[0]})`, body: target },
			{ method: "PATCH", url: `accounts(${all[1]})`, body: target },
		]);
		expect(progress).toHaveBeenLastCalledWith(expect.objectContaining({ completed: 3, total: 3, totalBatches: 2 }));
	});

	it("formats per-record Dataverse errors with the code (verified 0x80048408 message)", async () => {
		const [a, b] = ids(2);
		const message = `1 is not a valid status code for state code AccountState.Inactive on account with Id ${b}.`;
		const executeBatch = vi.fn(async () => [ok(), fail(400, "0x80048408", message)]);
		const result = await updateRecordStatesInBatches("accounts", [a, b], target, { executeBatch, wait: noWait });
		expect(result.failed).toBe(1);
		expect(result.errors[0]).toBe(`${b}: 0x80048408: ${message}`);
	});

	it("redoes a stopped batch one record at a time; a 404 still fails for a PATCH", async () => {
		const [a, b, c] = ids(3);
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) throw new Error("Malformed Dataverse batch response: invalid content ID correlation");
			if (requests[0].url.includes(b)) return [notFound(b)];
			return [ok()];
		});
		const result = await updateRecordStatesInBatches("accounts", [a, b, c], target, { executeBatch, wait: noWait });
		expect(executeBatch).toHaveBeenCalledTimes(4);
		expect(result.succeeded).toBe(2);
		expect(result.failed).toBe(1);
		// Known codes get their guidance first (formatDataverseError, #96)
		expect(result.errors[0]).toMatch(new RegExp(`^${b}: The record or object wasn.t found.* Details: 0x80040217: `));
	});

	it("stops when the connection changes and reports the rest as not changed (#66)", async () => {
		const all = ids(3);
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		const result = await updateRecordStatesInBatches("accounts", all, target, {
			executeBatch,
			batchSize: 1,
			shouldContinue: (() => {
				let n = 0;
				return () => ++n <= 2;
			})(),
			wait: noWait,
		});
		expect(result.succeeded).toBe(2);
		expect(result.errors).toEqual([`${all[2]}: ${STATE_CONNECTION_CHANGED_MESSAGE}`]);
	});
});
