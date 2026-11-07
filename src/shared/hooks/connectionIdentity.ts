/**
 * Which connection the tool is working against (#66).
 *
 * Since PPTB 1.2.7 a running tool's connection can change in place ("Change connection", or the
 * footer's connection slots), with no reload: the host re-sends the tool context and emits
 * connection:updated. Every Dataverse call then goes to the new environment. But
 * connection:updated is also broadcast to every tool when any connection is edited, activated or
 * has its token refreshed, so it isn't a "my connection changed" signal on its own: compare the
 * connection's identity (id + URL) before and after.
 */

/** The parts of a ToolBox connection that identify it */
export interface ConnectionLike {
	id?: string | null;
	url?: string | null;
}

/**
 * A stable key for a connection: its id and environment URL (lower case, no trailing slash).
 * "" for no connection (a connectionless launch, supported since PPTB 1.2.6).
 */
export function connectionKey(connection: ConnectionLike | null | undefined): string {
	if (!connection || (!connection.id && !connection.url)) return "";
	const url = (connection.url ?? "").trim().replace(/\/+$/, "").toLowerCase();
	return `${connection.id ?? ""}|${url}`;
}

/** The state that's scoped to one connection: remounted (and caches cleared) when it changes */
export interface ConnectionScope {
	/** Key of the connection the scope belongs to; undefined until the first read */
	key: string | undefined;
	/** Incremented on every change after the first read */
	session: number;
}

/**
 * The scope after reading the connection key. The first read only records the key: the tool
 * started on that connection, so nothing is stale. Any later different key is a new session.
 */
export function nextConnectionScope(scope: ConnectionScope, key: string | undefined): { scope: ConnectionScope; changed: boolean } {
	if (key === undefined || key === scope.key) return { scope, changed: false };
	if (scope.key === undefined) return { scope: { ...scope, key }, changed: false };
	return { scope: { key, session: scope.session + 1 }, changed: true };
}
