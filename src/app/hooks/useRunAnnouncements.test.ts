/**
 * What a run says to screen readers (#186 PR 3): "Running the query…", then the first page ("3 rows
 * loaded, more available. Ran in 120 ms."), nothing for a failed run (its bar speaks), Retrieve all
 * progress at most every PROGRESS_EVERY_MS and "…, all loaded." at the end, but not after Stop.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, render } from "../../shared/testing/ui";
import { AnnouncerProvider } from "../../shared/feedback/Announcer";
import { AnnouncerContext, FLUSH_MS } from "../../shared/feedback/useAnnouncer";
import { PROGRESS_EVERY_MS, useRunAnnouncements } from "./useRunAnnouncements";
import type { QueryExecution } from "./useQueryExecution";

type State = Pick<
	QueryExecution,
	"isExecuting" | "isRetrieveAllInProgress" | "queryResult" | "executeError" | "pagingNotice"
>;
const idle: State = {
	isExecuting: false,
	isRetrieveAllInProgress: false,
	queryResult: null,
	executeError: undefined,
	pagingNotice: null,
};
function Probe(state: State) {
	useRunAnnouncements(state);
	return null;
}
const show = (state: Partial<State>) =>
	render(createElement(AnnouncerProvider, null, createElement(Probe, { ...idle, ...state })));
const rows = (n: number) => Array.from({ length: n }, (_, i) => ({ id: i }));
const said = () => document.querySelector('[data-fxs-announcer="polite"]')?.textContent ?? "";
const flushed = () => act(() => new Promise<void>((resolve) => setTimeout(resolve, FLUSH_MS + 40)));

afterEach(() => {
	cleanup();
	vi.restoreAllMocks();
});

describe("useRunAnnouncements", () => {
	it("says the run started, then how many rows came, whether more are available, and the time", async () => {
		await show({});
		await show({ isExecuting: true });
		await flushed();
		expect(said()).toBe("Running the query…");
		await show({ queryResult: { rows: rows(3), moreRecords: true, executionTimeMs: 120 } });
		await flushed();
		expect(said()).toBe("3 rows loaded, more available. Ran in 120 ms.");
	});

	it("says all loaded, and one row, when that's so", async () => {
		await show({ isExecuting: true });
		await show({ queryResult: { rows: rows(1), moreRecords: false } });
		await flushed();
		expect(said()).toBe("1 row, all loaded.");
	});

	it("says nothing for a failed run, and doesn't read 'Running…' after it", async () => {
		await show({ isExecuting: true });
		await show({ executeError: "Dataverse said no" });
		await flushed();
		expect(said()).toBe("");
	});

	it("starts over after a run whose results were cleared (the entity changed)", async () => {
		await show({ isExecuting: true });
		await flushed();
		await show({});
		await show({ isExecuting: true });
		await flushed();
		await show({ queryResult: { rows: rows(2), moreRecords: false } });
		await flushed();
		expect(said()).toBe("2 rows, all loaded.");
		// And the second run said it started
		await show({ isExecuting: true });
		await flushed();
		expect(said()).toBe("Running the query…");
	});

	it("says the run started once, though the old results clear while it runs", async () => {
		const spoken = vi.fn();
		const showSpied = (state: Partial<State>) =>
			render(
				createElement(
					AnnouncerContext.Provider,
					{ value: spoken },
					createElement(Probe, { ...idle, ...state }),
				),
			);
		const before = { rows: rows(5), moreRecords: false };
		await showSpied({ queryResult: before });
		await showSpied({ isExecuting: true, queryResult: before });
		await showSpied({ isExecuting: true, queryResult: null });
		expect(spoken.mock.calls.map(([message]) => message)).toEqual(["Running the query…"]);
	});

	it("says Retrieve all's progress at most every so often, then all loaded", async () => {
		let now = 1_000_000;
		vi.spyOn(Date, "now").mockImplementation(() => now);
		await show({ isExecuting: true });
		await show({
			isRetrieveAllInProgress: true,
			queryResult: { rows: rows(50), moreRecords: true },
		});
		await flushed();
		expect(said()).toBe("50 rows loaded, loading the rest.");
		// A page soon after: nothing new
		now += PROGRESS_EVERY_MS / 2;
		await show({
			isRetrieveAllInProgress: true,
			queryResult: { rows: rows(100), moreRecords: true },
		});
		await flushed();
		expect(said()).toBe("50 rows loaded, loading the rest.");
		now += PROGRESS_EVERY_MS;
		await show({
			isRetrieveAllInProgress: true,
			queryResult: { rows: rows(1500), moreRecords: true },
		});
		await flushed();
		expect(said()).toBe("1,500 rows loaded so far.");
		await show({ queryResult: { rows: rows(1520), moreRecords: false } });
		await flushed();
		expect(said()).toBe("1,520 rows, all loaded.");
	});

	it("says all loaded even with the 'no paging cookie' note", async () => {
		await show({ isExecuting: true });
		await show({
			isRetrieveAllInProgress: true,
			queryResult: { rows: rows(50), moreRecords: true },
		});
		await flushed();
		await show({
			queryResult: { rows: rows(80), moreRecords: false },
			pagingNotice: { kind: "simple", message: "Paging without a paging cookie" },
		});
		await flushed();
		expect(said()).toBe("80 rows, all loaded.");
	});

	it("leaves Stop to its own bar: no 'all loaded'", async () => {
		await show({ isExecuting: true });
		await show({
			isRetrieveAllInProgress: true,
			queryResult: { rows: rows(50), moreRecords: true },
		});
		await flushed();
		await show({
			queryResult: { rows: rows(100), moreRecords: true },
			pagingNotice: { kind: "stopped", message: "" },
		});
		await flushed();
		expect(said()).toBe("50 rows loaded, loading the rest.");
	});
});
