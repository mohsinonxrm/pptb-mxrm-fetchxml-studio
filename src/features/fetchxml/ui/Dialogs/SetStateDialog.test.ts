/**
 * Activate / Deactivate confirmation (#189, rules from #69 and #143): names the environment and the
 * scope, says how to undo, plans the change from the entity's state model, and runs it once.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { toEntityStateModel } from "../../model/recordState";
import { button, cleanup, click, deferred, dialogInput, dialogText, flush, installHost, isDisabled, removeHost, render, scopeField, typeInto } from "../../../../shared/testing/ui";
import type { BatchDeleteResult } from "../../api/pptbClient";

const label = (text: string) => ({ UserLocalizedLabel: { Label: text } });
// Account's state model as the sandbox returned it (recordState.test.ts, 2026-10-17)
const accountModel = toEntityStateModel(
	[
		{ Value: 0, Label: label("Active"), InvariantName: "Active", DefaultStatus: 1 },
		{ Value: 1, Label: label("Inactive"), InvariantName: "Inactive", DefaultStatus: 2 },
	],
	[
		{ Value: 1, Label: label("Active"), State: 0, TransitionData: null },
		{ Value: 2, Label: label("Inactive"), State: 1, TransitionData: null },
	],
	false,
);

vi.mock("../../api/pptbClient", async (importOriginal) => ({
	...(await importOriginal<typeof import("../../api/pptbClient")>()),
	getEntityStateModel: vi.fn(async () => accountModel),
	getRecordStatuses: vi.fn(async () => undefined),
}));

// Imported after the mock is declared (vi.mock is hoisted, so this order is for readers)
const { SetStateDialog } = await import("./SetStateDialog");

const ids = (n: number) => Array.from({ length: n }, (_, i) => `00000000-0000-0000-0000-${String(i).padStart(12, "0")}`);

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

const done = (n: number): BatchDeleteResult => ({ succeeded: n, failed: 0, errors: [] }) as unknown as BatchDeleteResult;

describe("Activate / Deactivate dialog", () => {
	it("Deactivate names the environment and scope, says Activate undoes it, and runs once with Inactive", async () => {
		installHost({ connection: { name: "CRM279420" } });
		const run = deferred<BatchDeleteResult>();
		const onExecute = vi.fn((_target: unknown, _onProgress: unknown) => run.promise);
		await render(
			createElement(SetStateDialog, {
				open: true,
				action: "deactivate",
				recordIds: ids(2),
				entityLogicalName: "account",
				entityDisplayName: "Account",
				primaryIdAttribute: "accountid",
				onClose: () => {},
				onExecute,
			}),
		);
		await flush();
		const text = dialogText();
		expect(text).toContain("CRM279420");
		expect(scopeField("Scope")).toBe("2 records");
		expect(text).toContain("To undo, run Activate");

		const deactivate = () => button("Deactivate");
		expect(isDisabled(deactivate())).toBe(false);
		await click(deactivate());
		expect(onExecute).toHaveBeenCalledTimes(1);
		expect(onExecute.mock.calls[0][0]).toEqual({ statecode: 1, statuscode: 2 });
		// Running: the button is gone and Cancel is disabled
		expect(deactivate()).toBeUndefined();
		expect(isDisabled(button("Cancel"))).toBe(true);

		run.resolve(done(2));
		await flush();
		expect(onExecute).toHaveBeenCalledTimes(1);
	});

	it("Activate says Deactivate undoes it", async () => {
		installHost();
		await render(
			createElement(SetStateDialog, {
				open: true,
				action: "activate",
				recordIds: ids(1),
				entityLogicalName: "account",
				entityDisplayName: "Account",
				primaryIdAttribute: "accountid",
				onClose: () => {},
				onExecute: async () => done(1),
			}),
		);
		await flush();
		expect(dialogText()).toContain("To undo, run Deactivate");
		expect(isDisabled(button("Activate"))).toBe(false);
	});

	it("asks for the entity name before changing more than 100 records", async () => {
		installHost();
		await render(
			createElement(SetStateDialog, {
				open: true,
				action: "deactivate",
				recordIds: ids(101),
				entityLogicalName: "account",
				entityDisplayName: "Account",
				primaryIdAttribute: "accountid",
				onClose: () => {},
				onExecute: async () => done(101),
			}),
		);
		await flush();
		expect(isDisabled(button("Deactivate"))).toBe(true);
		await typeInto(dialogInput(), "account");
		expect(isDisabled(button("Deactivate"))).toBe(false);
	});
});
