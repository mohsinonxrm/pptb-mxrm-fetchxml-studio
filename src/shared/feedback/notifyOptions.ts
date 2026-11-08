/**
 * App feedback (#277): the pure part of the notification provider. Which timeout, politeness and
 * dedupe key a toast gets, so the rules are testable without rendering anything.
 */

export type NotifyIntent = "success" | "info" | "warning" | "error";

export interface NotifyInput {
	title: string;
	body?: string;
	/** One optional action (Undo, Try again, View in operations); it runs, then the toast goes */
	action?: { label: string; onClick: () => void };
	/**
	 * A second action (Copy FetchXML), shown after `action`; it runs, then the toast goes. Needs
	 * `action`: a secondary action without a primary one is a programming error and is ignored.
	 */
	secondaryAction?: { label: string; onClick: () => void };
	/** Toasts with the same key replace each other; defaults to intent, title and body */
	key?: string;
	/** Overrides the intent's default; "sticky" stays until dismissed */
	timeout?: number | "sticky";
}

export interface ResolvedOptions {
	/** Milliseconds, or -1 for a toast that stays until dismissed (Fluent's convention) */
	timeout: number;
	sticky: boolean;
	politeness: "polite";
	key: string;
}

const SUCCESS_MS = 4000;
const INFO_MS = 5000;
const WARNING_MS = 8000;
/** A toast with an action stays long enough to read it and press the action */
const ACTION_MIN_MS = 8000;

export function dedupeKey(intent: NotifyIntent, input: NotifyInput): string {
	return input.key ?? `${intent}|${input.title}|${input.body ?? ""}`;
}

export function resolveOptions(intent: NotifyIntent, input: NotifyInput): ResolvedOptions {
	const key = dedupeKey(intent, input);
	const requested: number | "sticky" =
		input.timeout ??
		(intent === "success"
			? SUCCESS_MS
			: intent === "info"
				? INFO_MS
				: intent === "warning"
					? WARNING_MS
					: "sticky");
	if (requested === "sticky") return { timeout: -1, sticky: true, politeness: "polite", key };
	const timeout = input.action ? Math.max(requested, ACTION_MIN_MS) : requested;
	return { timeout, sticky: false, politeness: "polite", key };
}
