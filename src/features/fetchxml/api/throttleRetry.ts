/**
 * Retry after Dataverse service protection errors (#96). Pure apart from the injected sleep.
 *
 * Microsoft: a 429 from the Web API carries a Retry-After header with the seconds to wait; clients
 * should wait that long and send the request again. Without it, the Learn sample waits 2^attempt
 * seconds. The PPTB v1.2.7 host doesn't retry 429 itself and doesn't pass Retry-After on, so the
 * exponential fallback is what normally applies; a Retry-After on the error is honoured when present.
 * https://learn.microsoft.com/power-apps/developer/data-platform/api-limits#retry-operations
 *
 * Only service protection errors are retried: Dataverse rejected those requests without running
 * them, so sending a delete or workflow again doesn't repeat it.
 */

import { isServiceProtectionError, parseDataverseError } from "./dataverseErrors";

export interface ThrottleRetryOptions {
	/** Retries after the first attempt (default 3) */
	maxRetries?: number;
	/** Longest wait before one retry (default 60 s). A longer Retry-After isn't retried early: the error is thrown. */
	maxDelayMs?: number;
	/** Wait before retry n is baseDelayMs × 2^(n-1) without a Retry-After (default 2 s: 2, 4, 8 s like Learn's sample) */
	baseDelayMs?: number;
	/** Which errors to retry (default: service protection errors) */
	shouldRetry?: (error: unknown) => boolean;
	/** Waits; injected by tests */
	sleep?: (ms: number) => Promise<void>;
	/** Called before each wait (debug logging, progress text) */
	onRetry?: (info: { attempt: number; delayMs: number; error: unknown }) => void;
	/**
	 * Checked after each wait, before the retry (#175): false stops with RetryStoppedError, e.g. when
	 * the tool's connection changed in place while waiting (connectionGuard, #66). A write retried
	 * after that would run in the new environment, where the same record IDs can exist.
	 */
	shouldContinue?: () => boolean | Promise<boolean>;
}

/** A retry wasn't sent because shouldContinue said stop (#175) */
export class RetryStoppedError extends Error {
	constructor(message = "Not sent: the tool's connection changed while waiting to retry") {
		super(message);
		this.name = "RetryStoppedError";
	}
}

export const DEFAULT_MAX_RETRIES = 3;
export const DEFAULT_MAX_DELAY_MS = 60_000;
export const DEFAULT_BASE_DELAY_MS = 2_000;

const defaultSleep = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

/**
 * How long to wait before retry `attempt` (1-based): the error's Retry-After when it has one,
 * otherwise exponential back-off. Undefined when Retry-After is longer than `maxDelayMs`.
 */
export function retryDelayMs(
	error: unknown,
	attempt: number,
	{
		maxDelayMs = DEFAULT_MAX_DELAY_MS,
		baseDelayMs = DEFAULT_BASE_DELAY_MS,
	}: ThrottleRetryOptions = {},
): number | undefined {
	const retryAfter = parseDataverseError(error).retryAfterSeconds;
	if (retryAfter !== undefined) {
		const ms = retryAfter * 1000;
		return ms <= maxDelayMs ? ms : undefined;
	}
	return Math.min(baseDelayMs * 2 ** (attempt - 1), maxDelayMs);
}

/**
 * Run `operation`, retrying it after service protection errors up to `maxRetries` times.
 * Any other error, or the last throttling error, is thrown unchanged.
 */
export async function withThrottleRetry<T>(
	operation: () => Promise<T>,
	options: ThrottleRetryOptions = {},
): Promise<T> {
	const {
		maxRetries = DEFAULT_MAX_RETRIES,
		shouldRetry = isServiceProtectionError,
		sleep = defaultSleep,
		onRetry,
		shouldContinue,
	} = options;
	for (let attempt = 1; ; attempt++) {
		try {
			return await operation();
		} catch (error) {
			if (attempt > maxRetries || !shouldRetry(error)) throw error;
			const delayMs = retryDelayMs(error, attempt, options);
			if (delayMs === undefined) throw error;
			onRetry?.({ attempt, delayMs, error });
			await sleep(delayMs);
			if (shouldContinue && !(await shouldContinue())) throw new RetryStoppedError();
		}
	}
}
