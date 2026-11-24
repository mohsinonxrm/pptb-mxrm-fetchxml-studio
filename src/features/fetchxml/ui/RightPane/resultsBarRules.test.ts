/**
 * Why each results command is off (#232): results without record IDs first, then the query's
 * shape, then the entity's rules, then the selection. The #174 / #177 / #143 wording is kept.
 */
import { describe, expect, it } from "vitest";
import { resultsBarReasons, type ResultsBarState } from "./resultsBarRules";

const state = (extra: Partial<ResultsBarState> = {}): ResultsBarState => ({
	selectedCount: 1,
	loadedCount: 10,
	canBulkDelete: true,
	...extra,
});

describe("resultsBarReasons", () => {
	it("has no reasons with one row selected", () => {
		expect(resultsBarReasons(state())).toEqual({});
	});

	it("asks for one row to open it or copy its URL", () => {
		expect(resultsBarReasons(state({ selectedCount: 0 }))).toMatchObject({
			open: "Select a row to open it",
			copyUrl: "Select a row to copy its URL",
		});
		expect(resultsBarReasons(state({ selectedCount: 2 }))).toMatchObject({
			open: "Select only one row to open it",
			copyUrl: "Select only one row to copy its URL",
		});
	});

	it("asks for rows to activate, deactivate, close or run a workflow on", () => {
		expect(resultsBarReasons(state({ selectedCount: 0 }))).toMatchObject({
			activate: "Select rows to activate them",
			deactivate: "Select rows to deactivate them",
			close: "Select rows to close them",
			runWorkflow: "Select rows to run a workflow on them",
		});
	});

	it("puts results without record IDs before everything else", () => {
		const noIds =
			"Aggregate query results don't contain record IDs, so record actions aren't available.";
		const reasons = resultsBarReasons(
			state({
				selectedCount: 0,
				recordActionsDisabledReason: noIds,
				deleteDisabledReason: "shape",
				workflowDisabledReason: "shape",
				deleteAllDisabledReason: "changed",
				activateDisabledReason: "entity",
				deactivateDisabledReason: "entity",
				closeDisabledReason: "entity",
			}),
		);
		expect(reasons).toMatchObject({
			open: noIds,
			copyUrl: noIds,
			activate: noIds,
			deactivate: noIds,
			close: noIds,
			delete: noIds,
			runWorkflow: noIds,
		});
	});

	it("puts the query's shape and the entity's rules before the selection", () => {
		const reasons = resultsBarReasons(
			state({
				selectedCount: 0,
				deleteDisabledReason: "rows repeat (delete)",
				workflowDisabledReason: "rows repeat (workflow)",
				deleteAllDisabledReason: "changed",
				activateDisabledReason: "no active state",
				deactivateDisabledReason: "use Close",
				closeDisabledReason: "rows repeat (close)",
			}),
		);
		expect(reasons).toMatchObject({
			delete: "rows repeat (delete)",
			runWorkflow: "rows repeat (workflow)",
			activate: "no active state",
			deactivate: "use Close",
			close: "rows repeat (close)",
		});
	});

	it("uses Delete all's own reason with no selection (#174)", () => {
		expect(
			resultsBarReasons(state({ selectedCount: 0, deleteAllDisabledReason: "Run it again" }))
				.delete,
		).toBe("Run it again");
		expect(resultsBarReasons(state({ selectedCount: 0 })).delete).toBeUndefined();
	});

	it("needs Bulk Delete for more than 100 selected (#177)", () => {
		expect(
			resultsBarReasons(state({ selectedCount: 101, canBulkDelete: false })).delete,
		).toMatch(/Bulk Delete/);
		expect(
			resultsBarReasons(state({ selectedCount: 101, canBulkDelete: true })).delete,
		).toBeUndefined();
		expect(
			resultsBarReasons(state({ selectedCount: 100, canBulkDelete: false })).delete,
		).toBeUndefined();
	});

	it("has nothing to copy without rows", () => {
		expect(resultsBarReasons(state({ loadedCount: 0 })).copy).toBe("There are no rows to copy");
	});

	it("while a run waits for its first page, the commands that change records wait too, before any other reason (#279)", () => {
		const reasons = resultsBarReasons(
			state({ isRunning: true, recordActionsDisabledReason: "No record IDs" }),
		);
		expect(reasons).toMatchObject({
			activate: "Wait for the query to finish",
			deactivate: "Wait for the query to finish",
			close: "Wait for the query to finish",
			delete: "Wait for the query to finish",
			runWorkflow: "Wait for the query to finish",
		});
		// Open, Copy URL and Copy only read: they keep their own reasons
		expect(reasons.open).toBe("No record IDs");
		expect(resultsBarReasons(state({ isRunning: true })).copy).toBeUndefined();
	});
});
