/**
 * PPTB Tool-to-Tool (T2T) Invocation helpers.
 *
 * Wraps the host invocation API (toolboxAPI.invocation) for both directions:
 *  - Caller: launching other tools, and discovering them by capability tag.
 *  - Callee: reading the inbound launch context and returning data.
 *
 * The invocation API is typed by @pptb/types (≥ 1.2.3-beta.0). We still access it through a
 * runtime guard, since an older host may not expose it (or its newer methods) at runtime.
 */

/** The host invocation API as typed on the global window.toolboxAPI. */
type ToolboxInvocationAPI = NonNullable<Window["toolboxAPI"]>["invocation"];

/**
 * The prefill payload FetchXML Studio sends when launching another tool via T2T.
 *
 * Shape is intentionally minimal — any callee can resolve additional context
 * (display name, entity set name, etc.) via its own dataverseAPI using entityLogicalName.
 */
export interface FetchXmlStudioT2TPrefill {
	/** Serialized FetchXML query string — the primary payload. */
	fetchXml: string;
	/** Root entity logical name, e.g. "account". */
	entityLogicalName: string;
}

/**
 * The launch-context prefill shape FetchXML Studio accepts when it is itself
 * invoked as a callee. Mirrors the prefill contract declared in pptb.config.json.
 *
 * The contract is deliberately permissive: a caller may pass a bare `fetchXml`
 * string, or a `viewRef` pointing at an existing saved/personal view. Note that
 * `entityLogicalName` is a hint only — the root entity is ALWAYS derived from the
 * FetchXML itself, since this tool is metadata-driven and a mismatched hint would
 * silently break every metadata lookup.
 */
export interface FetchXmlStudioLaunchPrefill {
	/** Serialized FetchXML query string. Root entity is parsed from this. */
	fetchXml?: string;
	/** Optional hint only — ignored for metadata resolution. */
	entityLogicalName?: string;
	/** Reference to an existing view; FetchXML is retrieved from the view record. */
	viewRef?: {
		id: string;
		entityLogicalName: "savedquery" | "userquery";
	};
}

/** Safely obtain the invocation API, or undefined when unavailable at runtime. */
function getInvocationAPI(): ToolboxInvocationAPI | undefined {
	if (typeof window === "undefined" || !window.toolboxAPI) return undefined;
	// Typed as always-present, but may be absent on older hosts — treat as optional.
	return (window.toolboxAPI as { invocation?: ToolboxInvocationAPI }).invocation;
}

/**
 * Returns true when the PPTB host exposes the T2T invocation API.
 * Use this to conditionally render T2T UI elements.
 */
export function isT2TSupported(): boolean {
	return getInvocationAPI() !== undefined;
}

/**
 * A tool discovered via capability search — the subset of the host Tool manifest we use.
 */
export interface DiscoveredTool {
	/** npm package id — passed as targetToolId to launchTool. */
	id: string;
	/** Human-readable display name (falls back to id). */
	name: string;
}

/** Our own package id, excluded from discovery results so we don't list ourselves. */
const SELF_TOOL_ID = "@mohsinonxrm/pptb-fetchxml-studio";

/**
 * Discover installed tools that accept FetchXML (capability tag "fetchxml"), excluding
 * ourselves. Returns [] when the host doesn't expose capability discovery (older hosts),
 * so callers degrade gracefully.
 */
export async function discoverFetchXmlTools(): Promise<DiscoveredTool[]> {
	const invocation = getInvocationAPI();
	if (!invocation?.findToolsByCapability) return [];

	try {
		const tools = (await invocation.findToolsByCapability("fetchxml")) as Array<{
			id?: unknown;
			name?: unknown;
		}>;
		return tools
			.filter((t): t is { id: string; name?: unknown } => {
				return typeof t?.id === "string" && t.id !== SELF_TOOL_ID;
			})
			.map((t) => ({
				id: t.id,
				name: typeof t.name === "string" && t.name.trim() !== "" ? t.name : t.id,
			}));
	} catch (error) {
		console.error("T2T discovery: findToolsByCapability('fetchxml') failed", error);
		return [];
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// Callee side — receiving a launch context and returning data to the caller
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Read the launch context for the current tool, if any.
 * Returns the prefill payload when this tool was opened via T2T, or null for a
 * standalone launch (or when the invocation API is unavailable).
 */
export async function getLaunchContext(): Promise<Record<string, unknown> | null> {
	const invocation = getInvocationAPI();
	if (!invocation) return null;
	return invocation.getLaunchContext();
}

/**
 * The action to take when consuming an inbound launch prefill:
 *  - "fetchxml": load this FetchXML into the builder (root entity is parsed from it)
 *  - "entity":   start a fresh query rooted at this entity (metadata-driven)
 *  - null:       nothing to prefill (standalone launch or empty payload)
 */
export type LaunchPrefillAction =
	| { kind: "fetchxml"; fetchXml: string }
	| { kind: "entity"; entityLogicalName: string }
	| null;

/**
 * Resolve what to prefill from a raw launch context, following the contract
 * declared in pptb.config.json. Resolution priority:
 *
 *   1. `fetchXml` string            → load it; root entity is derived from the FetchXML,
 *                                      and any `entityLogicalName` hint is ignored (a
 *                                      mismatch would break this metadata-driven tool).
 *   2. `viewRef`                    → retrieve the view's FetchXML, then as (1).
 *   3. `entityLogicalName` only     → no query to parse, so use the entity as the
 *                                      starting point for a fresh query.
 *
 * @returns The action to apply, or null when there is nothing to prefill.
 */
export async function resolveLaunchPrefill(
	context: Record<string, unknown> | null,
): Promise<LaunchPrefillAction> {
	if (!context) return null;
	const prefill = context as FetchXmlStudioLaunchPrefill;

	if (typeof prefill.fetchXml === "string" && prefill.fetchXml.trim() !== "") {
		return { kind: "fetchxml", fetchXml: prefill.fetchXml };
	}

	if (prefill.viewRef && typeof prefill.viewRef.id === "string") {
		const fetchXml = await resolveViewRefFetchXml(prefill.viewRef);
		return fetchXml ? { kind: "fetchxml", fetchXml } : null;
	}

	if (typeof prefill.entityLogicalName === "string" && prefill.entityLogicalName.trim() !== "") {
		return { kind: "entity", entityLogicalName: prefill.entityLogicalName.trim() };
	}

	return null;
}

/**
 * Retrieve a view's FetchXML by reference. Reuses the same dataverseAPI.queryData
 * route that the entity → view picker uses to read view definitions (savedquery /
 * userquery), filtered to a single record by id.
 */
async function resolveViewRefFetchXml(viewRef: {
	id: string;
	entityLogicalName: "savedquery" | "userquery";
}): Promise<string | null> {
	const dataverse = window.dataverseAPI;
	if (!dataverse) return null;

	const isPersonal = viewRef.entityLogicalName === "userquery";
	const setName = isPersonal ? "userqueries" : "savedqueries";
	const idField = isPersonal ? "userqueryid" : "savedqueryid";

	const query = `${setName}?$select=${idField},fetchxml&$filter=${idField} eq ${viewRef.id}`;
	const result = await dataverse.queryData(query);
	const record = result.value?.[0];
	const fetchXml = record?.fetchxml;

	return typeof fetchXml === "string" && fetchXml.trim() !== "" ? fetchXml : null;
}

/** Return data to the invoking tool. No-op when this tool was not launched via T2T. */
export async function returnDataToInvokingTool(data: Record<string, unknown>): Promise<void> {
	const invocation = getInvocationAPI();
	if (!invocation) return;
	await invocation.returnData(data);
}

/** Convenience helper for returning the current FetchXML document to the caller. */
export async function returnFetchXmlToInvokingTool(fetchXml: string): Promise<void> {
	await returnDataToInvokingTool({ fetchXml });
}

/**
 * Launch a target PPTB tool with the current FetchXML query as prefill data.
 * Automatically forwards the active Dataverse connection via primaryConnectionId.
 *
 * @returns The data returned by the callee via returnData(), or null if the callee
 *          closed without returning data.
 * @throws  If the target tool is not installed or the launch fails.
 */
export async function sendFetchXmlToTool(
	targetToolId: string,
	prefill: FetchXmlStudioT2TPrefill,
): Promise<unknown> {
	const invocation = getInvocationAPI();
	if (!invocation) {
		throw new Error(
			"PPTB invocation API is not available. Upgrade to PPTB host ≥ 1.2.2 to use Tool-to-Tool invocation.",
		);
	}

	const connection = await window.toolboxAPI.connections.getActiveConnection();
	return invocation.launchTool(targetToolId, prefill as unknown as Record<string, unknown>, {
		primaryConnectionId: connection?.id ?? null,
		// "Send to Tool" is a one-way handoff — suppress the callee's "Return to FXS" banner.
		noReturn: true,
	});
}
