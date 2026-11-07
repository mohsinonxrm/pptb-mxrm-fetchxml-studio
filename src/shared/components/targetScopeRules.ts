/**
 * What a write operation's confirmation says about its target (#69): the environment and the scope.
 * Pure, so the wording and the type-to-confirm rule are unit-tested.
 *
 * PPTB Marketplace policy for tools that change data: confirmations name the target environment and
 * the scope (which table, how many records). https://docs.powerplatformtoolbox.com/policies/marketplace
 */

/** More records than this need the entity logical name typed before the action runs (#69) */
export const TYPE_TO_CONFIRM_RECORD_THRESHOLD = 100;

/** What the operation applies to */
export type TargetScopeKind =
	/** These exact records (the selected rows) */
	| { kind: "records"; count: number }
	/**
	 * Every record the query matches (Bulk Delete with no selection). The total isn't known: only the
	 * rows loaded in the grid, and whether Dataverse said more records match.
	 */
	| { kind: "allMatching"; loadedCount?: number; moreRecords?: boolean }
	/** Anything else, already worded (e.g. the view a Save writes) */
	| { kind: "text"; text: string };

/** PPTB connection environment types: "Dev" | "Test" | "UAT" | "Production" (@pptb/types 1.2.7) */
export const isProductionEnvironment = (environment: string | undefined): boolean =>
	environment?.trim().toLowerCase() === "production";

const records = (count: number) => (count === 1 ? "1 record" : `${count.toLocaleString()} records`);

/**
 * The scope as one line. Never claims a total it doesn't know: with no selection, the count is
 * what the grid loaded, and says so.
 */
export function describeScope(scope: TargetScopeKind): string {
	switch (scope.kind) {
		case "records":
			return records(scope.count);
		case "text":
			return scope.text;
		case "allMatching": {
			const base = "Every record the query matches";
			const loaded = scope.loadedCount;
			if (!loaded) return base;
			if (scope.moreRecords === true) return `${base} (more than the ${records(loaded)} loaded)`;
			if (scope.moreRecords === false) return `${base} (${records(loaded)} when the query last ran)`;
			return `${base} (${records(loaded)} loaded; there may be more)`;
		}
	}
}

/**
 * Whether the user types the entity logical name before the action (#69): on a Production
 * connection, for more than TYPE_TO_CONFIRM_RECORD_THRESHOLD records, and when the total is unknown.
 */
export function requiresTypedConfirmation(environment: string | undefined, scope: TargetScopeKind): boolean {
	if (isProductionEnvironment(environment)) return true;
	if (scope.kind === "records") return scope.count > TYPE_TO_CONFIRM_RECORD_THRESHOLD;
	return scope.kind === "allMatching";
}

/** Whether the typed text confirms: the expected word, ignoring case and surrounding spaces */
export const isTypedConfirmation = (typed: string, expected: string): boolean =>
	typed.trim().toLowerCase() === expected.trim().toLowerCase() && expected.trim() !== "";
