/**
 * Close dialogs (#189, rules from #69 and #149): name the environment and the scope, say how to
 * undo (Activate; the close activities stay), and send the close message once.
 */
import { createElement } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CLOSE_ACTIONS, type CloseMessageAction } from "../../model/closeActions";
import { toEntityStateModel } from "../../model/recordState";
import { button, cleanup, click, deferred, dialogInput, dialogText, flush, installHost, isDisabled, removeHost, render, scopeField, typeInto } from "../../../../shared/testing/ui";
import type { BatchDeleteResult } from "../../api/pptbClient";

const label = (text: string) => ({ UserLocalizedLabel: { Label: text } });
// Opportunity's state model: Open, Won, Lost, with their default status reasons
const opportunityModel = toEntityStateModel(
	[
		{ Value: 0, Label: label("Open"), InvariantName: "Open", DefaultStatus: 1 },
		{ Value: 1, Label: label("Won"), InvariantName: "Won", DefaultStatus: 3 },
		{ Value: 2, Label: label("Lost"), InvariantName: "Lost", DefaultStatus: 4 },
	],
	[
		{ Value: 1, Label: label("In Progress"), State: 0, TransitionData: null },
		{ Value: 3, Label: label("Won"), State: 1, TransitionData: null },
		{ Value: 4, Label: label("Canceled"), State: 2, TransitionData: null },
		{ Value: 5, Label: label("Out-Sold"), State: 2, TransitionData: null },
	],
	false,
);

vi.mock("../../api/pptbClient", async (importOriginal) => ({
	...(await importOriginal<typeof import("../../api/pptbClient")>()),
	getEntityStateModel: vi.fn(async () => opportunityModel),
	searchCompetitors: vi.fn(async () => []),
}));

const { CloseRecordsDialog } = await import("./CloseRecordsDialog");

const win = CLOSE_ACTIONS.opportunity.find((a) => a.key === "win") as CloseMessageAction;
const ids = (n: number) => Array.from({ length: n }, (_, i) => `00000000-0000-0000-0000-${String(i).padStart(12, "0")}`);
const done = (n: number): BatchDeleteResult => ({ succeeded: n, failed: 0, errors: [] }) as unknown as BatchDeleteResult;

beforeEach(() => {
	vi.spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
	cleanup();
	removeHost();
	vi.restoreAllMocks();
});

describe("Close as won dialog", () => {
	it("names the environment and scope, says Activate undoes it, and closes once as Won", async () => {
		installHost({ connection: { name: "CRM279420" } });
		const run = deferred<BatchDeleteResult>();
		const onExecute = vi.fn((_values: unknown, _onProgress: unknown) => run.promise);
		await render(
			createElement(CloseRecordsDialog, {
				open: true,
				action: win,
				recordIds: ids(2),
				entityLogicalName: "opportunity",
				entityDisplayName: "Opportunity",
				onClose: () => {},
				onExecute,
			}),
		);
		await flush();
		const text = dialogText();
		expect(text).toContain("CRM279420");
		expect(scopeField("Scope")).toBe("2 records");
		expect(text).toContain("To undo, run Activate on the records this changes. The close activities stay.");

		const closeButton = () => button("Close as won");
		expect(isDisabled(closeButton())).toBe(false);
		await click(closeButton());
		expect(onExecute).toHaveBeenCalledTimes(1);
		expect(onExecute.mock.calls[0][0]).toMatchObject({ statuscode: 3 });
		expect(closeButton()).toBeUndefined();
		expect(isDisabled(button("Cancel"))).toBe(true);

		run.resolve(done(2));
		await flush();
		expect(onExecute).toHaveBeenCalledTimes(1);
	});

	it("asks for the entity name before closing more than 100 records", async () => {
		installHost();
		await render(
			createElement(CloseRecordsDialog, {
				open: true,
				action: win,
				recordIds: ids(120),
				entityLogicalName: "opportunity",
				entityDisplayName: "Opportunity",
				onClose: () => {},
				onExecute: async () => done(120),
			}),
		);
		await flush();
		expect(isDisabled(button("Close as won"))).toBe(true);
		await typeInto(dialogInput(), "opportunity");
		expect(isDisabled(button("Close as won"))).toBe(false);
	});
});
