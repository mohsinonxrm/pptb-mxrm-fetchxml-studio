/**
 * PPTB host API adapter (#70): one place for feature detection and fallbacks, so FetchXML Studio
 * degrades gracefully on hosts that don't expose a newer API.
 *
 * Source for every API here: the @pptb/types 1.2.7 typings (toolboxAPI.d.ts, dataverseAPI.d.ts)
 * and README. https://docs.powerplatformtoolbox.com/tool-development/api-reference/toolbox-api
 */

import { connectionKey } from "../../../shared/hooks/connectionIdentity";

type ToolboxAPI = Window["toolboxAPI"];
export type Connection = NonNullable<Awaited<ReturnType<ToolboxAPI["connections"]["getActiveConnection"]>>>;

/** window.toolboxAPI, or undefined outside the ToolBox (dev server, tests) */
const toolbox = (): Partial<ToolboxAPI> | undefined =>
	typeof window === "undefined" ? undefined : (window.toolboxAPI as Partial<ToolboxAPI> | undefined);

/**
 * The tool's primary connection (slot 0). PPTB ≥ 1.2.7: connections.getConnection(0), the
 * slot-aware API; older hosts: getActiveConnection(). Null when the tool has no connection.
 */
export async function getPrimaryConnection(): Promise<Connection | null> {
	const connections = toolbox()?.connections;
	if (!connections) return null;
	if (typeof connections.getConnection === "function") return (await connections.getConnection(0)) ?? null;
	if (typeof connections.getActiveConnection === "function") return (await connections.getActiveConnection()) ?? null;
	return null;
}

/**
 * For multi-request write operations (#66): the returned check says whether the tool is still on
 * the connection it started on. Since PPTB 1.2.7 the connection can change in place, and every
 * later request goes to the new environment, where the same record IDs can exist (copied
 * environments). No guard (always true) when the connection can't be read at the start; a check
 * that can't read it says no.
 */
export async function connectionGuard(): Promise<() => Promise<boolean>> {
	let start: string;
	try {
		start = connectionKey(await getPrimaryConnection());
	} catch {
		return async () => true;
	}
	return async () => {
		try {
			return connectionKey(await getPrimaryConnection()) === start;
		} catch {
			return false;
		}
	};
}

/**
 * Opens a URL in the browser profile of the tool's connection, already signed in to the right
 * tenant (utils.openInConnectionBrowser; https/http only). Falls back to window.open.
 */
export async function openUrl(url: string): Promise<void> {
	const open = toolbox()?.utils?.openInConnectionBrowser;
	if (typeof open === "function") {
		try {
			await open(url);
			return;
		} catch (error) {
			console.warn("openInConnectionBrowser failed, using window.open:", error);
		}
	}
	window.open(url, "_blank", "noopener");
}

/**
 * Writes text to the clipboard. The ToolBox's utils.copyToClipboard comes first: it works in the
 * tool's webview without the focus and permission quirks of navigator.clipboard (#70); the
 * browser clipboard is the fallback. Returns whether it worked.
 */
export async function copyText(text: string): Promise<boolean> {
	const copy = toolbox()?.utils?.copyToClipboard;
	if (typeof copy === "function") {
		try {
			await copy(text);
			return true;
		} catch (error) {
			console.warn("copyToClipboard failed, using navigator.clipboard:", error);
		}
	}
	try {
		await navigator.clipboard.writeText(text);
		return true;
	} catch {
		return false;
	}
}

// Prevent-close holds are counted, so overlapping operations (a Retrieve All while a delete runs)
// release the host's warning only when the last one ends.
let closeHolds = 0;

/**
 * Asks the host to warn before the tool tab or PPTB closes while a long-running operation is in
 * progress (toolboxAPI.preventClose, PPTB ≥ 1.2.6). Returns the release function; calling it more
 * than once is harmless. A no-op on hosts without the API.
 */
export function holdClose(): () => void {
	const api = toolbox();
	if (typeof api?.preventClose !== "function" || typeof api.releasePreventClose !== "function") {
		return () => {};
	}
	if (closeHolds++ === 0) {
		api.preventClose().catch((error: unknown) => console.warn("preventClose failed:", error));
	}
	let released = false;
	return () => {
		if (released) return;
		released = true;
		if (--closeHolds === 0) {
			api.releasePreventClose!().catch((error: unknown) => console.warn("releasePreventClose failed:", error));
		}
	};
}

/** Runs an operation while holding the host's close warning (see holdClose). */
export async function whileHoldingClose<T>(operation: () => Promise<T>): Promise<T> {
	const release = holdClose();
	try {
		return await operation();
	} finally {
		release();
	}
}

/** Whether the host exposes dataverseAPI.executeBatch (PPTB ≥ 1.2.7) */
export function canExecuteBatch(): boolean {
	return typeof window !== "undefined" && typeof window.dataverseAPI?.executeBatch === "function";
}
