/**
 * Why each results command can't be used right now (#232). Every disabled command carries its own
 * reason, on the bar and in More, instead of the "Record actions unavailable" and "Delete (all)
 * unavailable" info buttons. Commands the user lacks the privilege for are hidden instead (owner
 * decision 1), so they aren't here.
 *
 * Order matters: the first reason that applies is the one shown. Results without record IDs come
 * first (nothing can act on them), then the query's shape, then the entity's own rules, then the
 * selection.
 */
import { deleteSelectionUnavailableReason } from "../../model/deleteScope";

/** While a run waits for its first page, the commands that change records or export wait too (#279) */
export const QUERY_RUNNING_REASON = "Wait for the query to finish";

export interface ResultsBarState {
	/** A run is waiting for its first page (owner, #279: not Retrieve all's later pages) */
	isRunning?: boolean;
	selectedCount: number;
	/** Rows the grid has (or the quick search matches) */
	loadedCount: number;
	/** No record IDs in the results (aggregate, distinct, no primary key) */
	recordActionsDisabledReason?: string;
	/** Delete is off for this query's shape (aggregate, or rows that repeat) */
	deleteDisabledReason?: string;
	/** Run workflow is off for this query's shape */
	workflowDisabledReason?: string;
	/** Delete all (no selection) is off: the query changed since it ran, or it hasn't run (#174) */
	deleteAllDisabledReason?: string;
	/** More than 100 selected needs Bulk Delete (#177) */
	canBulkDelete: boolean;
	activateDisabledReason?: string;
	deactivateDisabledReason?: string;
	closeDisabledReason?: string;
}

export interface ResultsBarReasons {
	open?: string;
	copyUrl?: string;
	activate?: string;
	deactivate?: string;
	close?: string;
	delete?: string;
	runWorkflow?: string;
	copy?: string;
}

export function resultsBarReasons(state: ResultsBarState): ResultsBarReasons {
	const n = state.selectedCount;
	// A run first: the rows on screen are about to be replaced
	const running = state.isRunning ? QUERY_RUNNING_REASON : undefined;
	const noIds = state.recordActionsDisabledReason;
	const pick = (verb: string) => (n === 0 ? `Select rows to ${verb}` : undefined);
	const one = (none: string, many: string) =>
		noIds ?? (n === 0 ? none : n > 1 ? many : undefined);
	return {
		open: one("Select a row to open it", "Select only one row to open it"),
		copyUrl: one("Select a row to copy its URL", "Select only one row to copy its URL"),
		activate: running ?? noIds ?? state.activateDisabledReason ?? pick("activate them"),
		deactivate: running ?? noIds ?? state.deactivateDisabledReason ?? pick("deactivate them"),
		close: running ?? noIds ?? state.closeDisabledReason ?? pick("close them"),
		// No selection is Delete all, with its own rules (#174); a large selection needs Bulk Delete (#177)
		delete:
			running ??
			noIds ??
			state.deleteDisabledReason ??
			(n === 0
				? state.deleteAllDisabledReason
				: deleteSelectionUnavailableReason(n, state.canBulkDelete)),
		runWorkflow:
			running ?? noIds ?? state.workflowDisabledReason ?? pick("run a workflow on them"),
		copy: state.loadedCount === 0 ? "There are no rows to copy" : undefined,
	};
}
