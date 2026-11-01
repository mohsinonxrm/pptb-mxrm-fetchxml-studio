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
import { formatDataverseError, isServiceProtectionError } from "./dataverseErrors";
import {
	DEFAULT_MAX_RETRIES,
	LONG_RUNNING_RETRY,
	RetryStoppedError,
	retryDelayMs,
	type ThrottleRetryOptions,
} from "./throttleRetry";
import { statePatchBody, type StateTarget } from "../model/recordState";

type HostExecuteBatch = Window["dataverseAPI"]["executeBatch"];
export type BatchRequest = Parameters<HostExecuteBatch>[0][number];
export type BatchResult = Awaited<ReturnType<HostExecuteBatch>>[number];
export type ExecuteBatch = (requests: BatchRequest[]) => Promise<BatchResult[]>;

/** Records per $batch. Small enough for steady progress and short requests; Learn allows 1,000. */
export const DELETE_BATCH_SIZE = 25;


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
/** Service protection error codes: number of requests, execution time, concurrent requests */
const SERVICE_PROTECTION_CODES = new Set(["0x80072322", "0x80072321", "0x80072326"]);
/** An operation the batch answered with a service protection error: Dataverse didn't run it */
const isThrottledResult = (result: BatchResult) =>
	result.status === 429 || SERVICE_PROTECTION_CODES.has(batchResultError(result).code?.toLowerCase() ?? "");
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
	/** Injectable for tests: the jitter's random number (LONG_RUNNING_RETRY) */
	random?: () => number;
	/**
	 * Back-off after service protection errors (#183): LONG_RUNNING_RETRY by default; a request the
	 * user is waiting on (one record) passes {} for the short default policy
	 */
	retry?: ThrottleRetryOptions;
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
	/**
	 * Before a stopped batch is redone one record per batch: the records the stopped batch already
	 * completed, which aren't sent again (close messages aren't repeatable, #149)
	 */
	completedInStoppedBatch?: (recordIds: string[]) => Promise<Set<string>>;
}

/**
 * Sends one batch, retrying while Dataverse throttles it (LONG_RUNNING_RETRY, #183): the whole
 * batch when it's rejected, and otherwise only the operations it answered with a service protection
 * error, which Dataverse didn't run. One retry budget covers both, so the waits stay within the
 * policy. A throttled operation's Retry-After is honoured; a longer one than the policy allows
 * isn't waited for.
 *
 * After each wait, `shouldContinue` is asked again (#175): the connection can change in place
 * during a wait, and resent requests would run in the new environment. If it says stop before
 * anything ran, RetryStoppedError is thrown; once some operations ran, the ones not sent again are
 * undefined in the result. `onWait` gets the wait in seconds, for the progress text.
 */
async function sendWithRetry(
	requests: BatchRequest[],
	execute: ExecuteBatch,
	{ wait, shouldContinue, onWait, random, policy = LONG_RUNNING_RETRY }: {
		wait: (ms: number) => Promise<void>;
		shouldContinue: () => boolean | Promise<boolean>;
		onWait?: (seconds: number) => void;
		random?: () => number;
		policy?: ThrottleRetryOptions;
	},
): Promise<(BatchResult | undefined)[]> {
	const retry: ThrottleRetryOptions = { ...policy, random };
	const maxRetries = retry.maxRetries ?? DEFAULT_MAX_RETRIES;
	// Undefined until the batch has answered; then each operation's latest result
	let results: (BatchResult | undefined)[] | undefined;
	let pending = requests.map((_, i) => i);
	for (let attempt = 1; ; attempt++) {
		let delayMs: number | undefined;
		try {
			const response = await execute(pending.map((i) => requests[i]));
			if (!results) {
				// A batch that stopped early is redone one record per batch by the caller
				if (response.length !== requests.length) return response;
				results = [...response];
			} else {
				if (response.length !== pending.length) return results;
				pending.forEach((index, k) => (results![index] = response[k]));
			}
			pending = pending.filter((i) => isThrottledResult(results![i]!));
			if (pending.length === 0 || attempt > maxRetries) return results;
			delayMs = retryDelayMs(results[pending[0]], attempt, retry);
			if (delayMs === undefined) return results;
		} catch (error) {
			const retryable = attempt <= maxRetries && isServiceProtectionError(error);
			delayMs = retryable ? retryDelayMs(error, attempt, retry) : undefined;
			if (delayMs === undefined) {
				// A failed resend leaves the throttled operations with their result (not run)
				if (results) return results;
				throw error;
			}
		}
		onWait?.(Math.ceil(delayMs / 1000));
		await wait(delayMs);
		if (!(await shouldContinue())) {
			if (!results) throw new RetryStoppedError();
			return results.map((r, i) => (pending.includes(i) ? undefined : r));
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

/** Why records weren't closed after shouldContinue said stop (#149) */
export const CLOSE_CONNECTION_CHANGED_MESSAGE = "Not closed: the tool's connection changed while closing";

/**
 * Close messages (#149): POST the action (WinOpportunity, CloseIncident, ...) once per record in
 * $batch requests, with the per-record errors, progress and connection guard delete has. Unlike
 * delete and PATCH, a close isn't repeatable (a second WinOpportunity answers 0x80040515 "The
 * opportunity is already closed"; verified on a sandbox), so when a batch stops early
 * `completedInStoppedBatch` names the records it already closed and they aren't sent again.
 * Verified on a sandbox (2026-10-17): a $batch of actions stops at the first failure and the
 * earlier ones stay done.
 */
export function closeRecordsInBatches(
	recordIds: string[],
	toRequest: (recordId: string) => BatchRequest,
	options: RecordBatchOptions & { completedInStoppedBatch?: (recordIds: string[]) => Promise<Set<string>> },
): Promise<BatchDeleteResult> {
	return runRecordBatches(recordIds, options, {
		toRequest,
		stoppedMessage: CLOSE_CONNECTION_CHANGED_MESSAGE,
		formatError: (result) => formatDataverseError(result.body && typeof result.body === "object" ? result.body : new Error(batchResultError(result).message)),
		completedInStoppedBatch: options.completedInStoppedBatch,
	});
}

/** Sends one request per record in $batch requests of `batchSize` (see the file comment) */
async function runRecordBatches(
	recordIds: string[],
	{ executeBatch, onProgress, batchSize = DELETE_BATCH_SIZE, shouldContinue = () => true, wait = sleep, random, retry }: RecordBatchOptions,
	{ toRequest, stoppedMessage, retriedSuccess = () => false, formatError = (r) => batchResultError(r).message, completedInStoppedBatch }: RecordOperation,
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
	// The last progress reported, repeated with the wait while Dataverse throttles (#183)
	let lastProgress: BatchDeleteProgress = {
		completed: 0,
		total: recordIds.length,
		succeeded: 0,
		failed: 0,
		batchesCompleted: 0,
		totalBatches: chunks.length,
	};
	const send = (requests: BatchRequest[]) =>
		sendWithRetry(requests, executeBatch, {
			wait,
			shouldContinue,
			random,
			policy: retry,
			onWait: (seconds) => onProgress?.({ ...lastProgress, throttleWaitSeconds: seconds }),
		});
	let stopped = false;
	const stop = (ids: string[]) => ids.forEach((id) => record(id, new Error(stoppedMessage), false));
	for (let index = 0; index < chunks.length; index++) {
		const chunk = chunks[index];
		const started = Date.now();
		const requests = chunk.map(toRequest);

		let results: (BatchResult | undefined)[] | undefined;
		let failedEarly = false;
		if (stopped || !(await shouldContinue())) {
			stopped = true;
			stop(chunk);
		} else {
			try {
				results = await send(requests);
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
			chunk.forEach((id, i) => {
				const outcome = results![i];
				if (outcome) record(id, outcome, false);
				else {
					// Stopped while waiting to resend it (#175): it didn't run
					stopped = true;
					record(id, new Error(stoppedMessage), false);
				}
			});
		} else if (failedEarly || results) {
			let completedIds = new Set<string>();
			if (completedInStoppedBatch) {
				try {
					completedIds = await completedInStoppedBatch(chunk);
				} catch (error) {
					console.warn(`Batch ${index + 1}: couldn't tell which records it completed, redoing all:`, error);
				}
			}
			result.succeeded += chunk.filter((id) => completedIds.has(id)).length;
			const redo = chunk.filter((id) => !completedIds.has(id));
			for (const [i, id] of redo.entries()) {
				if (!(await shouldContinue())) {
					stopped = true;
					stop(redo.slice(i));
					break;
				}
				try {
					const sent = await send([toRequest(id)]);
					if (sent.length > 0 && !sent[0]) throw new RetryStoppedError();
					record(id, sent[0] ?? new Error("No response for this record"), true);
				} catch (error) {
					if (error instanceof RetryStoppedError) {
						// Stopped while waiting to retry this record (#175): it and the rest weren't sent
						stopped = true;
						stop(redo.slice(i));
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
		lastProgress = {
			completed,
			total: recordIds.length,
			succeeded: result.succeeded,
			failed: result.failed,
			batchesCompleted: index + 1,
			totalBatches: chunks.length,
			estimatedSecondsRemaining: chunkTimes.length >= 2 ? Math.ceil((average * remaining) / 1000) : undefined,
		};
		onProgress?.(lastProgress);
	}
	return result;
}
