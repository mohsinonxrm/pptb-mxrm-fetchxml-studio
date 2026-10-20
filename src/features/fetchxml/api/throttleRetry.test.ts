import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { RetryStoppedError, retryDelayMs, withThrottleRetry } from "./throttleRetry";

// #96: back-off after service protection errors, with fake timers.

const throttled = () =>
	new Error(
		"0x80072322: Number of requests exceeded the limit of 6000 over time window of 300 seconds.",
	);

beforeEach(() => {
	vi.useFakeTimers();
});

afterEach(() => {
	vi.useRealTimers();
});

describe("retryDelayMs (#96)", () => {
	it("backs off exponentially without Retry-After (2, 4, 8 s), capped", () => {
		expect(retryDelayMs(throttled(), 1)).toBe(2000);
		expect(retryDelayMs(throttled(), 2)).toBe(4000);
		expect(retryDelayMs(throttled(), 3)).toBe(8000);
		expect(retryDelayMs(throttled(), 10, { maxDelayMs: 30_000 })).toBe(30_000);
	});

	it("waits Retry-After when the error carries it, and doesn't retry early past the cap", () => {
		const withRetryAfter = (seconds: number) => ({
			error: { code: "0x80072322", message: "Number of requests exceeded the limit." },
			headers: { "retry-after": String(seconds) },
		});
		expect(retryDelayMs(withRetryAfter(5), 1)).toBe(5000);
		expect(retryDelayMs(withRetryAfter(300), 1)).toBeUndefined();
	});
});

describe("withThrottleRetry (#96)", () => {
	it("retries a throttled operation after the back-off and returns its result", async () => {
		const operation = vi.fn().mockRejectedValueOnce(throttled()).mockResolvedValueOnce("ok");
		const promise = withThrottleRetry(operation);
		await vi.advanceTimersByTimeAsync(1999);
		expect(operation).toHaveBeenCalledTimes(1);
		await vi.advanceTimersByTimeAsync(1);
		await expect(promise).resolves.toBe("ok");
		expect(operation).toHaveBeenCalledTimes(2);
	});

	it("gives up after maxRetries and throws the last error unchanged", async () => {
		const error = throttled();
		const operation = vi.fn().mockRejectedValue(error);
		const onRetry = vi.fn();
		const promise = withThrottleRetry(operation, { maxRetries: 2, onRetry });
		const settled = expect(promise).rejects.toBe(error);
		await vi.advanceTimersByTimeAsync(2000 + 4000);
		await settled;
		expect(operation).toHaveBeenCalledTimes(3);
		expect(onRetry.mock.calls.map(([info]) => info.delayMs)).toEqual([2000, 4000]);
	});

	it("doesn't retry other errors", async () => {
		const error = new Error(
			"0x80040220: Principal user is missing prvDeleteAccount privilege.",
		);
		const operation = vi.fn().mockRejectedValue(error);
		await expect(withThrottleRetry(operation)).rejects.toBe(error);
		expect(operation).toHaveBeenCalledTimes(1);
	});

	it("doesn't retry after shouldContinue says stop during the wait (#175)", async () => {
		const sleep = vi.fn().mockResolvedValue(undefined);
		const operation = vi.fn().mockRejectedValue(throttled());
		await expect(withThrottleRetry(operation, { sleep, shouldContinue: () => false })).rejects.toBeInstanceOf(RetryStoppedError);
		expect(sleep).toHaveBeenCalledTimes(1);
		expect(operation).toHaveBeenCalledTimes(1); // waited, then not resent
	});

	it("uses an injected sleep", async () => {
		const sleep = vi.fn().mockResolvedValue(undefined);
		const operation = vi
			.fn()
			.mockRejectedValueOnce(new Error("HTTP 429"))
			.mockResolvedValueOnce(1);
		await expect(withThrottleRetry(operation, { sleep })).resolves.toBe(1);
		expect(sleep).toHaveBeenCalledWith(2000);
	});
});
