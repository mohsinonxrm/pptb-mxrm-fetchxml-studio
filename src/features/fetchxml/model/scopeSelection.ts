/**
 * The publisher / solution selection remembered per environment (#25).
 */

/** One environment's remembered selection */
export interface RememberedScope {
	publisherIds: string[];
	solutionIds: string[];
}

/** Every environment's selection, keyed by environmentKey() */
export type RememberedScopes = Record<string, RememberedScope>;

/** toolboxAPI.settings key (namespaced like displaySettings) */
export const REMEMBERED_SCOPES_KEY = "fetchxml-studio:scopeSelection";

/** An environment URL as a key: lower case, no trailing slash; "" without one */
export function environmentKey(url: string | null | undefined): string {
	return (url ?? "").trim().replace(/\/+$/, "").toLowerCase();
}

const idList = (value: unknown): string[] => (Array.isArray(value) ? value.filter((v): v is string => typeof v === "string" && !!v) : []);

/** The stored map, or an empty one when the stored value isn't one */
export function readRememberedScopes(stored: unknown): RememberedScopes {
	if (!stored || typeof stored !== "object" || Array.isArray(stored)) return {};
	const scopes: RememberedScopes = {};
	for (const [key, value] of Object.entries(stored as Record<string, unknown>)) {
		if (!value || typeof value !== "object") continue;
		const entry = value as Record<string, unknown>;
		scopes[key] = { publisherIds: idList(entry.publisherIds), solutionIds: idList(entry.solutionIds) };
	}
	return scopes;
}

/** The environment's remembered selection, if any */
export function rememberedScopeFor(stored: unknown, envKey: string): RememberedScope | undefined {
	if (!envKey) return undefined;
	return readRememberedScopes(stored)[envKey];
}

/** The map with the environment's selection replaced (removed when empty) */
export function withRememberedScope(stored: unknown, envKey: string, scope: RememberedScope): RememberedScopes {
	const scopes = readRememberedScopes(stored);
	if (!envKey) return scopes;
	if (scope.publisherIds.length === 0 && scope.solutionIds.length === 0) delete scopes[envKey];
	else scopes[envKey] = { publisherIds: [...scope.publisherIds], solutionIds: [...scope.solutionIds] };
	return scopes;
}

/**
 * The part of a remembered selection that still exists. Ids no longer there (deleted publisher or
 * solution, another environment's copy) are dropped silently.
 */
export function restorableIds(remembered: readonly string[], available: Iterable<string>): string[] {
	const ids = new Set(available);
	return remembered.filter((id) => ids.has(id));
}
