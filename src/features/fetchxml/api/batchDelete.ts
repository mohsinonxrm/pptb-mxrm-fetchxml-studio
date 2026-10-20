/**
 * Record delete through the host's dataverseAPI.executeBatch (PPTB ≥ 1.2.7) with the metadata
 * EntitySetName (#61). dataverseAPI.delete(entityLogicalName, id) builds the URL with the host's
 * English pluralizer, which is wrong for ~10% of custom tables (verified live, #61); executeBatch
 * takes explicit relative URLs.
 *
 * Microsoft Learn, Execute batch operations using the Web API:
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/execute-batch-operations-using-web-api
 * - "Batch requests can contain up to 1,000 individual requests."
 * - Without `Prefer: odata.continue-on-error`, the batch stops at the first failing request. We
 *   don't send that header: any additional header makes PPTB ask the user for consent first.
 *
 * Host behaviour (PPTB 1.2.7 dataverseBatch.ts): executeBatch returns one result per operation,
 * but rejects when a batch stopped early (fewer results than operations). A one-operation batch
 * always returns its result. So a chunk that rejects is redone one record per batch: records the
 * stopped batch already deleted then answer 404 "Does Not Exist", which counts as deleted.
 *
 * The same runner sends Activate / Deactivate PATCHes (#143, updateRecordStatesInBatches). A PATCH
 * of statecode / statuscode to the state a record already has succeeds (verified on a sandbox,
 * 2026-10-17), so redoing a stopped chunk one record per batch is safe there too.
 */

import type { BatchDeleteProgress, BatchDeleteResult } from "./pptbClient";
import { formatDataverseError } from "./dataverseErrors";
import { RetryStoppedError } from "./throttleRetry";
import { statePatchBody, type StateTarget } from "../model/recordState";

type HostExecuteBatch = Window["dataverseAPI"]["executeBatch"];
export type BatchRequest = Parameters<HostExecuteBatch>[0][number];
export type BatchResult = Awaited<ReturnType<HostExecuteBatch>>[number];
export type ExecuteBatch = (requests: BatchRequest[]) => Promise<BatchResult[]>;

/** Records per $batch. Small enough for steady progress and short requests; Learn allows 1,000. */
export const DELETE_BATCH_SIZE = 25;

/** Service protection errors (429): number of requests, execution time, concurrent requests */
const THROTTLE_PATTERN = /\b429\b|0x80072322|0x80072321|0x80072326/i;
/** Waits before retrying a throttled batch. The host doesn't pass on Retry-After from a rejected batch. */
const THROTTLE_WAITS_MS = [5_000, 15_000, 30_000];

/** `${EntitySetName}(${id})`, the Web API URL of a record (relative, as executeBatch requires) */
export const recordUrl = (entitySetName: string, recordId: string) => `${entitySetName}(${recordId.replace(/[{}]/g, "")})`;

/** The Dataverse error code and message of a failed batch operation */
export function batchResultError(result: BatchResult): { code?: string; message: string } {
	const body = result.body as { error?: { code?: string; message?: string } } | string | undefined;
	if (body && typeof body === "object" && body.error) {
		return { code: body.error.code, message: body.error.message || `HTTP ${result.status}` };
	}
	const text = typeof body === "string" && body.trim() ? `: ${body.trim()}` : "";
	return { message: `HTTP ${result.status}${result.statusText ? ` ${result.statusText}` : ""}${text}` };
}

const isSuccess = (result: BatchResult) => result.status >= 200 && result.status < 300;

/**
 * The record doesn't exist (404, ObjectDoesNotExist 0x80040217). Not a 404 for an unknown entity
 * set segment (0x80060888): that one is a real failure.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/reference/web-service-error-codes
 */
export function isRecordNotFound(result: BatchResult): boolean {
	if (result.status !== 404) return false;
	const { code, message } = batchResultError(result);
	if (code?.toLowerCase() === "0x80040217") return true;
	return !/0x80060888|Resource not found for the segment/i.test(`${code} ${message}`) && /does not exist/i.test(message);
}

const errorMessage = (error: unknown) => (error instanceof Error ? error.message : String(error));
const isThrottled = (error: unknown) => THROTTLE_PATTERN.test(errorMessage(error));
const sleep = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

export interface RecordBatchOptions {
	executeBatch: ExecuteBatch;
	onProgress?: (progress: BatchDeleteProgress) => void;
	batchSize?: number;
	/**
	 * Checked before every request; false stops the delete and reports the remaining records as
	 * not deleted (host.connectionGuard: the tool's connection changed, #66)
	 */
	shouldContinue?: () => boolean | Promise<boolean>;
	/** Injectable for tests */
	wait?: (ms: number) => Promise<void>;
}

export type BatchDeleteOptions = RecordBatchOptions;

/** Why records weren't deleted after shouldContinue said stop */
export const CONNECTION_CHANGED_MESSAGE = "Not deleted: the tool's connection changed during the delete";

/** Why records weren't changed after shouldContinue said stop (#143) */
export const STATE_CONNECTION_CHANGED_MESSAGE = "Not changed: the tool's connection changed during the update";

/** What differs between the operations the runner sends */
interface RecordOperation {
	/** The request for one record */
	toRequest: (recordId: string) => BatchRequest;
	/** Why the remaining records weren't processed after shouldContinue said stop */
	stoppedMessage: string;
	/** Whether a failed result of a retried (one-record) batch still counts as done */
	retriedSuccess?: (result: BatchResult) => boolean;
	/** The text for a failed result */
	formatError?: (result: BatchResult) => string;
}

/**
 * Sends one batch, retrying while Dataverse throttles it. After each wait, `shouldContinue` is asked
 * again (#175): the connection can change in place during a 5–30 s wait, and the resent batch would
 * then run in the new environment.
 */
async function sendWithRetry(
	requests: BatchRequest[],
	execute: ExecuteBatch,
	wait: (ms: number) => Promise<void>,
	shouldContinue: () => boolean | Promise<boolean>,
): Promise<BatchResult[]> {
	for (let attempt = 0; ; attempt++) {
		try {
			return await execute(requests);
		} catch (error) {
			if (!isThrottled(error) || attempt >= THROTTLE_WAITS_MS.length) throw error;
			await wait(THROTTLE_WAITS_MS[attempt]);
			if (!(await shouldContinue())) throw new RetryStoppedError();
		}
	}
}

/**
 * Deletes records in $batch requests of `batchSize`, reporting per-record errors and progress
 * (same shape as the parallel-request path it replaces).
 */
export function deleteRecordsInBatches(entitySetName: string, recordIds: string[], options: BatchDeleteOptions): Promise<BatchDeleteResult> {
	return runRecordBatches(recordIds, options, {
		toRequest: (id) => ({ method: "DELETE", url: recordUrl(entitySetName, id) }),
		stoppedMessage: CONNECTION_CHANGED_MESSAGE,
		retriedSuccess: isRecordNotFound,
	});
}

/**
 * Activate / Deactivate (#143): PATCH `{ statecode, statuscode }` to each record in $batch requests,
 * with the same per-record errors, progress and connection guard as delete. Errors are formatted
 * with formatDataverseError (#96).
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/update-delete-entities-using-web-api#basic-update
 */
export function updateRecordStatesInBatches(
	entitySetName: string,
	recordIds: string[],
	target: StateTarget,
	options: RecordBatchOptions,
): Promise<BatchDeleteResult> {
	return runRecordBatches(recordIds, options, {
		toRequest: (id) => buildStatePatchRequest(entitySetName, id, target),
		stoppedMessage: STATE_CONNECTION_CHANGED_MESSAGE,
		formatError: (result) => formatDataverseError(result.body && typeof result.body === "object" ? result.body : new Error(batchResultError(result).message)),
	});
}

/** The batch request that sets a record's state (#143): PATCH {EntitySetName}(id) { statecode, statuscode } */
export const buildStatePatchRequest = (entitySetName: string, recordId: string, target: StateTarget): BatchRequest => ({
	method: "PATCH",
	url: recordUrl(entitySetName, recordId),
	body: statePatchBody(target),
});

/** Sends one request per record in $batch requests of `batchSize` (see the file comment) */
async function runRecordBatches(
	recordIds: string[],
	{ executeBatch, onProgress, batchSize = DELETE_BATCH_SIZE, shouldContinue = () => true, wait = sleep }: RecordBatchOptions,
	{ toRequest, stoppedMessage, retriedSuccess = () => false, formatError = (r) => batchResultError(r).message }: RecordOperation,
): Promise<BatchDeleteResult> {
	const result: BatchDeleteResult = { succeeded: 0, failed: 0, errors: [] };
	const size = Math.max(1, Math.min(batchSize, 1000));
	const chunks: string[][] = [];
	for (let i = 0; i < recordIds.length; i += size) chunks.push(recordIds.slice(i, i + size));

	const record = (recordId: string, outcome: BatchResult | Error, retried: boolean) => {
		if (outcome instanceof Error) {
			result.failed++;
			result.errors.push(`${recordId}: ${outcome.message}`);
		} else if (isSuccess(outcome) || (retried && retriedSuccess(outcome))) {
			result.succeeded++;
		} else {
			result.failed++;
			result.errors.push(`${recordId}: ${formatError(outcome)}`);
		}
	};

	const chunkTimes: number[] = [];
	let completed = 0;
	let stopped = false;
	const stop = (ids: string[]) => ids.forEach((id) => record(id, new Error(stoppedMessage), false));
	for (let index = 0; index < chunks.length; index++) {
		const chunk = chunks[index];
		const started = Date.now();
		const requests = chunk.map(toRequest);

		let results: BatchResult[] | undefined;
		let failedEarly = false;
		if (stopped || !(await shouldContinue())) {
			stopped = true;
			stop(chunk);
		} else {
			try {
				results = await sendWithRetry(requests, executeBatch, wait, shouldContinue);
			} catch (error) {
				if (error instanceof RetryStoppedError) {
					// Stopped while waiting to retry (#175): nothing in this chunk was sent
					stopped = true;
					stop(chunk);
				} else if (chunk.length === 1) {
					// One record per batch tells which records failed (see the file comment). A lone
					// record gets the error itself: there's nothing to tell apart.
					record(chunk[0], new Error(errorMessage(error)), false);
				} else {
					failedEarly = true;
					console.warn(`Batch ${index + 1} stopped early, retrying its records one by one:`, error);
				}
			}
		}

		if (results && results.length === chunk.length) {
			chunk.forEach((id, i) => record(id, results![i], false));
		} else if (failedEarly || results) {
			for (const [i, id] of chunk.entries()) {
				if (!(await shouldContinue())) {
					stopped = true;
					stop(chunk.slice(i));
					break;
				}
				try {
					const [single] = await sendWithRetry([toRequest(id)], executeBatch, wait, shouldContinue);
					record(id, single ?? new Error("No response for this record"), true);
				} catch (error) {
					if (error instanceof RetryStoppedError) {
						// Stopped while waiting to retry this record (#175): it and the rest weren't sent
						stopped = true;
						stop(chunk.slice(i));
						break;
					}
					record(id, new Error(errorMessage(error)), true);
				}
			}
		}

		completed += chunk.length;
		chunkTimes.push(Date.now() - started);
		const remaining = chunks.length - (index + 1);
		const average = chunkTimes.reduce((a, b) => a + b, 0) / chunkTimes.length;
		onProgress?.({
			completed,
			total: recordIds.length,
			succeeded: result.succeeded,
			failed: result.failed,
			batchesCompleted: index + 1,
			totalBatches: chunks.length,
			estimatedSecondsRemaining: chunkTimes.length >= 2 ? Math.ceil((average * remaining) / 1000) : undefined,
		});
	}
	return result;
}
