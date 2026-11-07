/**
 * Per-connection memo for Dataverse calls that don't change during a session (#71): WhoAmI and
 * privilege checks. Concurrent callers share one in-flight request, so startup sends each call
 * once instead of once per component (AppShell, EntitySelector, the record actions, Save view).
 *
 * The answers belong to the connected user and environment: call resetDataverseSessionCache()
 * when the active connection changes (or is removed), so the next call asks Dataverse again.
 * Failed calls aren't kept, so a later call retries them.
 */

const entries = new Map<string, Promise<unknown>>();

/**
 * The memoized result of `load` for `key` in this session. A rejected promise, or a value that
 * `keep` refuses (e.g. null from a WhoAmI that failed), is dropped so the next call loads again.
 */
export function memoizeForSession<T>(
	key: string,
	load: () => Promise<T>,
	keep: (value: T) => boolean = () => true,
): Promise<T> {
	const existing = entries.get(key) as Promise<T> | undefined;
	if (existing) return existing;
	const promise = load();
	entries.set(key, promise);
	const drop = () => {
		// Only if it's still this request: a reset (and a newer request) may have replaced it
		if (entries.get(key) === promise) entries.delete(key);
	};
	promise.then((value) => {
		if (!keep(value)) drop();
	}, drop);
	return promise;
}

/**
 * Forget every memoized WhoAmI and privilege answer (#71). Call it when the PPTB active connection
 * changes or is deleted; requests still in flight for the old connection aren't kept.
 */
export function resetDataverseSessionCache(): void {
	entries.clear();
}
