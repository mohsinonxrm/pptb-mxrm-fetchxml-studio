/**
 * The FetchXML editor's bars (#277, owner decision D3 (a), 2026-11-16): the editor owns every
 * message about its FetchXML, one bar per condition, in one group at the top of the FetchXML tab.
 * Validate, Parse to tree and the editor-mode checks of Execute, Save view and Send to tool all
 * report here, so the same problem is never shown twice. A check replaces the bars for what it
 * checked; a clean check removes them. #287 turns this group into the Problems panel.
 */
import type { ParseIssueItem } from "./ParseIssueList";

/**
 * - `error`: the XML or the FetchXML is invalid, or Dataverse is known to reject it
 * - `warning`: Dataverse might reject it
 * - `tree-changed`: the tree changed while the editor was being edited
 * - `notes`: settings Dataverse ignores, and other notes
 */
export type EditorProblemKey = "error" | "warning" | "tree-changed" | "notes";

export interface EditorProblem {
	key: EditorProblemKey;
	intent: "error" | "warning" | "info";
	title: string;
	/** Text, or a list of findings rendered as bullets */
	body: string | ParseIssueItem[];
	/** The Advisories tab lists one of these findings: offer "View all in Advisories" (#97) */
	advisories?: boolean;
	/** One action of its own, e.g. Execute anyway */
	action?: { label: string; onClick: () => void };
}

export const SYNTAX_ERROR_TITLE = "Syntax error";
export const INVALID_TITLE = "Invalid FetchXML";
export const MIGHT_REJECT_TITLE = "Dataverse might reject this query";
export const NOTES_TITLE = "Notes";

/** Most severe first: Fluent orders a group by severity */
const ORDER: EditorProblemKey[] = ["error", "warning", "tree-changed", "notes"];

/** What Validate and Parse to tree check: everything about the FetchXML */
export const FULL_CHECK: EditorProblemKey[] = ["error", "warning", "notes"];

export const syntaxProblem = (message: string): EditorProblem => ({
	key: "error",
	intent: "error",
	title: SYNTAX_ERROR_TITLE,
	body: message,
});

export const invalidProblem = (items: ParseIssueItem[], advisories?: boolean): EditorProblem => ({
	key: "error",
	intent: "error",
	title: INVALID_TITLE,
	body: items,
	advisories,
});

export const mightRejectProblem = (
	items: ParseIssueItem[],
	options: { advisories?: boolean; action?: EditorProblem["action"] } = {},
): EditorProblem => ({
	key: "warning",
	intent: "warning",
	title: MIGHT_REJECT_TITLE,
	body: items,
	...options,
});

export const notesProblem = (items: ParseIssueItem[], advisories?: boolean): EditorProblem => ({
	key: "notes",
	intent: "info",
	title: NOTES_TITLE,
	body: items,
	advisories,
});

/**
 * The bars after a check: the check's findings replace the bars for the keys it `checked` (a key
 * it checked and found nothing for is removed); other bars stay. Sorted most severe first.
 */
export function applyCheck(
	current: EditorProblem[],
	checked: EditorProblemKey[],
	found: EditorProblem[],
): EditorProblem[] {
	const replaced = new Set<EditorProblemKey>([...checked, ...found.map((p) => p.key)]);
	const next = [...current.filter((p) => !replaced.has(p.key)), ...found];
	return next.sort((a, b) => ORDER.indexOf(a.key) - ORDER.indexOf(b.key));
}
