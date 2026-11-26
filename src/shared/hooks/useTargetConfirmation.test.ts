/**
 * What a write confirmation knows before and after TargetScope reports (#279 review): never
 * confirmed before its first report, always with a reason when not confirmed, and not confirmed
 * again once TargetScope goes (a dialog that opens again reads the connection again).
 */
import { createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, installHost, removeHost, render } from "../testing/ui";
import {
	READING_CONNECTION,
	useTargetConfirmation,
	type TargetConfirmation,
} from "./useTargetConfirmation";
import { TargetScope } from "../components/TargetScope";

afterEach(() => {
	cleanup();
	removeHost();
});

let latest: TargetConfirmation;
function Probe() {
	latest = useTargetConfirmation();
	return null;
}

describe("useTargetConfirmation", () => {
	it("isn't confirmed before TargetScope reports, and says why", async () => {
		await render(createElement(Probe));
		expect(latest.confirmed).toBe(false);
		expect(latest.reason).toBe(READING_CONNECTION);
	});
});

describe("TargetScope", () => {
	it("says it's no longer confirmed when it goes", async () => {
		installHost();
		const onConfirmationChange = vi.fn();
		await render(
			createElement(TargetScope, {
				entityLogicalName: "contact",
				scope: { kind: "records", count: 1 },
				onConfirmationChange,
			}),
		);
		onConfirmationChange.mockClear();
		cleanup();
		expect(onConfirmationChange).toHaveBeenLastCalledWith(false, READING_CONNECTION);
	});
});
