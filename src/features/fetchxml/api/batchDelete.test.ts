import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
	CONNECTION_CHANGED_MESSAGE,
	CLOSE_CONNECTION_CHANGED_MESSAGE,
	STATE_CONNECTION_CHANGED_MESSAGE,
	closeRecordsInBatches,
	runWorkflowInBatches,
	WORKFLOW_UNKNOWN_MESSAGE,
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

	it("waits up to about 5 minutes for a throttled batch, and reports each wait (#183)", async () => {
		const wait = vi.fn((_ms: number) => Promise.resolve());
		const progress = vi.fn();
		const throttled = new Error("0x80072322: Number of requests exceeded the limit of 6000 over time window of 300 seconds.");
		const executeBatch = vi.fn<(requests: BatchRequest[]) => Promise<BatchResult[]>>();
		for (let i = 0; i < 6; i++) executeBatch.mockRejectedValueOnce(throttled);
		executeBatch.mockImplementation(async (requests) => requests.map(ok));
		// random 0.5 is no jitter
		const result = await deleteRecordsInBatches("accounts", ids(2), { executeBatch, wait, onProgress: progress, random: () => 0.5 });
		expect(wait.mock.calls.map(([ms]) => ms)).toEqual([5_000, 10_000, 20_000, 40_000, 80_000, 120_000]);
		expect(result.succeeded).toBe(2);
		const waits = progress.mock.calls.map(([p]) => p.throttleWaitSeconds).filter((s) => s !== undefined);
		expect(waits).toEqual([5, 10, 20, 40, 80, 120]);
		// The progress after the batch has no wait
		const last = progress.mock.calls.at(-1)?.[0];
		expect(last).toMatchObject({ completed: 2, succeeded: 2 });
		expect(last.throttleWaitSeconds).toBeUndefined();
	});

	it("resends only the operations a batch answered with a service protection error (#183)", async () => {
		const [a, b, c] = ids(3);
		const wait = vi.fn(noWait);
		const sent: string[][] = [];
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			sent.push(requests.map((r) => r.url));
			// First try: b throttled inside the batch; a and c deleted
			return sent.length === 1 ? [ok(), fail(429, "0x80072322", "Number of requests exceeded the limit"), ok()] : requests.map(ok);
		});
		const result = await deleteRecordsInBatches("accounts", [a, b, c], { executeBatch, wait, random: () => 0.5 });
		expect(sent).toEqual([[`accounts(${a})`, `accounts(${b})`, `accounts(${c})`], [`accounts(${b})`]]);
		expect(wait).toHaveBeenCalledTimes(1);
		expect(result).toMatchObject({ succeeded: 3, failed: 0 });
	});

	it("reports an operation still throttled after every retry as failed, without resending the others", async () => {
		const [a, b] = ids(2);
		const wait = vi.fn(noWait);
		const executeBatch = vi.fn(async (requests: BatchRequest[]) =>
			requests.map((r) => (r.url.includes(b) ? fail(429, "0x80072322", "Number of requests exceeded the limit") : ok())),
		);
		const result = await deleteRecordsInBatches("accounts", [a, b], { executeBatch, wait, random: () => 0.5 });
		expect(wait).toHaveBeenCalledTimes(6);
		expect(executeBatch.mock.calls.slice(1).every(([requests]) => requests.length === 1)).toBe(true);
		expect(result.succeeded).toBe(1);
		expect(result.failed).toBe(1);
		expect(result.errors[0]).toContain(b);
	});

	it("one retry budget covers rejected resends and throttled operations together (#183 review)", async () => {
		const [a, b] = ids(2);
		const wait = vi.fn(noWait);
		const throttled = new Error("0x80072322: Number of requests exceeded the limit");
		let calls = 0;
		// b is throttled inside the batch, then every resend alternates a rejection and a 429 again
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			calls++;
			if (calls === 1) return [ok(), fail(429, "0x80072322", "throttled")];
			if (calls % 2 === 0) throw throttled;
			return requests.map(() => fail(429, "0x80072322", "throttled"));
		});
		const result = await deleteRecordsInBatches("accounts", [a, b], { executeBatch, wait, random: () => 0.5 });
		expect(wait).toHaveBeenCalledTimes(6);
		expect(result).toMatchObject({ succeeded: 1, failed: 1 });
	});

	it("honours a throttled operation's Retry-After (#183 review)", async () => {
		const [a, b] = ids(2);
		const wait = vi.fn((_ms: number) => Promise.resolve());
		let calls = 0;
		const executeBatch = vi.fn(async (requests: BatchRequest[]) =>
			++calls === 1
				? [ok(), { ...fail(429, "0x80072322", "throttled"), headers: { "retry-after": "40" } }]
				: requests.map(ok),
		);
		const result = await deleteRecordsInBatches("accounts", [a, b], { executeBatch, wait, random: () => 0.5 });
		expect(wait.mock.calls.map(([ms]) => ms)).toEqual([40_000]);
		expect(result.succeeded).toBe(2);
	});

	it("an operation not resent because the connection changed is reported as not deleted, the rest as done (#175, #183 review)", async () => {
		const [a, b] = ids(2);
		let checks = 0;
		const executeBatch = vi.fn(async () => [ok(), fail(429, "0x80072322", "throttled")]);
		// On the connection for the batch, changed by the time the wait is over
		const result = await deleteRecordsInBatches("accounts", [a, b], { executeBatch, wait: noWait, shouldContinue: () => ++checks === 1 });
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(result.succeeded).toBe(1);
		expect(result.failed).toBe(1);
		expect(result.errors).toEqual([`${b}: ${CONNECTION_CHANGED_MESSAGE}`]);
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

describe("closeRecordsInBatches", () => {
	// #149: close messages are POSTs to the action, one per record, and not repeatable
	const toRequest = (id: string): BatchRequest => ({ method: "POST", url: "WinOpportunity", body: { OpportunityClose: { id }, Status: 3 } });
	const alreadyClosed = fail(400, "0x80040515", "The opportunity is already closed.");

	it("sends one POST per record in batches", async () => {
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		const result = await closeRecordsInBatches(ids(3), toRequest, { executeBatch, batchSize: 2, wait: noWait });
		expect(executeBatch).toHaveBeenCalledTimes(2);
		expect(executeBatch.mock.calls[0][0][0]).toEqual({ method: "POST", url: "WinOpportunity", body: { OpportunityClose: { id: ids(3)[0] }, Status: 3 } });
		expect(result).toEqual({ succeeded: 3, failed: 0, errors: [] });
	});

	it("doesn't resend records a stopped batch already closed", async () => {
		const [a, b, c] = ids(3);
		const sent: string[] = [];
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) throw new Error("Batch stopped early");
			const id = (requests[0].body as { OpportunityClose: { id: string } }).OpportunityClose.id;
			sent.push(id);
			return [id === b ? alreadyClosed : ok()];
		});
		const completedInStoppedBatch = vi.fn(async () => new Set([a]));
		const result = await closeRecordsInBatches([a, b, c], toRequest, { executeBatch, wait: noWait, completedInStoppedBatch });
		expect(completedInStoppedBatch).toHaveBeenCalledWith([a, b, c]);
		expect(sent).toEqual([b, c]);
		expect(result.succeeded).toBe(2);
		expect(result.failed).toBe(1);
		expect(result.errors[0]).toMatch(new RegExp(`^${b}: .*The opportunity is already closed`));
	});

	it("a stop while redoing reports only the records not yet redone, not the closed ones (#183 review)", async () => {
		const [a, b, c] = ids(3);
		let checks = 0;
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) throw new Error("Batch stopped early");
			// b is throttled; the connection changes while waiting to resend it
			throw new Error("0x80072322: Number of requests exceeded the limit");
		});
		// a was closed by the stopped batch, so only b and c are redone
		const completedInStoppedBatch = vi.fn(async () => new Set([a]));
		const result = await closeRecordsInBatches([a, b, c], toRequest, {
			executeBatch,
			wait: noWait,
			completedInStoppedBatch,
			shouldContinue: () => ++checks <= 2,
		});
		expect(result.succeeded + result.failed).toBe(3);
		expect(result.succeeded).toBe(1);
		expect(result.errors.map((e) => e.split(":")[0])).toEqual([b, c]);
		expect(result.errors.every((e) => e.endsWith(CLOSE_CONNECTION_CHANGED_MESSAGE))).toBe(true);
	});

	it("redoes the whole stopped batch when it can't tell what was closed", async () => {
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) throw new Error("Batch stopped early");
			return [ok()];
		});
		const result = await closeRecordsInBatches(ids(2), toRequest, {
			executeBatch,
			wait: noWait,
			completedInStoppedBatch: () => Promise.reject(new Error("read failed")),
		});
		expect(executeBatch).toHaveBeenCalledTimes(3);
		expect(result.succeeded).toBe(2);
	});

	it("stops when the connection changes and reports the rest as not closed (#66)", async () => {
		const all = ids(2);
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		const result = await closeRecordsInBatches(all, toRequest, {
			executeBatch,
			batchSize: 1,
			shouldContinue: (() => {
				let n = 0;
				return () => ++n <= 1;
			})(),
			wait: noWait,
		});
		expect(result.succeeded).toBe(1);
		expect(result.errors).toEqual([`${all[1]}: ${CLOSE_CONNECTION_CHANGED_MESSAGE}`]);
	});
});

describe("runWorkflowInBatches (#264, background workflows)", () => {
	const workflowId = "{11111111-2222-3333-4444-555555555555}";
	const entityIdOf = (request: BatchRequest) => (request.body as { EntityId: string }).EntityId;
	/** System jobs per record: `jobs` changes as the workflow "runs" */
	const jobStore = (initial: Record<string, string[]> = {}) => {
		const jobs = new Map(Object.entries(initial).map(([id, list]) => [id, new Set(list)]));
		return {
			start: (id: string) => {
				if (!jobs.has(id)) jobs.set(id, new Set());
				jobs.get(id)!.add(`job-${id}-${jobs.get(id)!.size}`);
			},
			jobsFor: vi.fn(async (chunk: string[]) => new Map(chunk.filter((id) => jobs.has(id)).map((id) => [id, new Set(jobs.get(id))]))),
		};
	};

	it("POSTs ExecuteWorkflow bound to the workflow once per record, in batches", async () => {
		const store = jobStore();
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		const result = await runWorkflowInBatches(workflowId, ids(3), { executeBatch, batchSize: 2, wait: noWait, jobsFor: store.jobsFor });
		expect(executeBatch).toHaveBeenCalledTimes(2);
		expect(executeBatch.mock.calls[0][0][0]).toEqual({
			method: "POST",
			url: "workflows(11111111-2222-3333-4444-555555555555)/Microsoft.Dynamics.CRM.ExecuteWorkflow",
			body: { EntityId: ids(3)[0] },
		});
		expect(result).toEqual({ succeeded: 3, failed: 0, errors: [] });
		// A snapshot of the jobs before each batch, for a batch that stops early
		expect(store.jobsFor).toHaveBeenCalledTimes(2);
	});

	it("doesn't resend records whose workflow started in a stopped batch; earlier runs don't count", async () => {
		const [a, b, c] = ids(3);
		// c ran this workflow before: an old job that must not count as started now
		const store = jobStore({ [c]: ["old-job"] });
		const sent: string[] = [];
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) {
				// The batch ran a, then b failed, and it stopped: c never ran
				store.start(a);
				throw new Error("Batch stopped early");
			}
			const id = entityIdOf(requests[0]);
			sent.push(id);
			store.start(id);
			return [id === b ? fail(400, "0x80040265", "The workflow failed for this record") : ok()];
		});
		const result = await runWorkflowInBatches(workflowId, [a, b, c], { executeBatch, wait: noWait, jobsFor: store.jobsFor });
		expect(sent).toEqual([b, c]);
		expect(result.succeeded).toBe(2);
		expect(result.errors).toEqual([expect.stringMatching(new RegExp(`^${b}: `))]);
	});

	it("sends nothing again when the jobs can't be read: a workflow mustn't run twice", async () => {
		const [a, b] = ids(2);
		let reads = 0;
		const jobsFor = vi.fn(async () => {
			if (++reads > 1) throw new Error("Principal user is missing prvReadAsyncOperation privilege");
			return new Map<string, Set<string>>();
		});
		const executeBatch = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) throw new Error("Batch stopped early");
			return [ok()];
		});
		const result = await runWorkflowInBatches(workflowId, [a, b], { executeBatch, wait: noWait, jobsFor });
		expect(executeBatch).toHaveBeenCalledTimes(1);
		expect(result).toEqual({ succeeded: 0, failed: 2, errors: [`${a}: ${WORKFLOW_UNKNOWN_MESSAGE}`, `${b}: ${WORKFLOW_UNKNOWN_MESSAGE}`] });
	});

	it("without a snapshot from before the batch, a stopped batch is unknown too; a batch that completes is fine", async () => {
		const [a, b] = ids(2);
		const jobsFor = vi.fn(async () => {
			throw new Error("Network error");
		});
		const stops = vi.fn(async (requests: BatchRequest[]) => {
			if (requests.length > 1) throw new Error("Batch stopped early");
			return [ok()];
		});
		expect((await runWorkflowInBatches(workflowId, [a, b], { executeBatch: stops, wait: noWait, jobsFor })).failed).toBe(2);
		const completes = vi.fn(async (requests: BatchRequest[]) => requests.map(ok));
		expect(await runWorkflowInBatches(workflowId, [a, b], { executeBatch: completes, wait: noWait, jobsFor })).toEqual({
			succeeded: 2,
			failed: 0,
			errors: [],
		});
	});
});
