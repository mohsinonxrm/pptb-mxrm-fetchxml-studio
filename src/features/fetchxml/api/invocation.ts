/**
 * PPTB Tool-to-Tool (T2T) Invocation – Caller API
 *
 * Handles launching other PPTB tools and passing FetchXML query context as prefill data.
 *
 * The invocation API (toolboxAPI.invocation) is introduced in PPTB host ≥ 1.2.2 and is not
 * yet reflected in the shipped @pptb/types declarations, so we define a local interface and
 * access it via a typed cast.
 */

/**
 * Runtime invocation API exposed by the PPTB host on toolboxAPI.invocation.
 * Mirrors the InvocationAPI described in the PPTB Inter-Tool Invocation documentation.
 */
interface InvocationAPI {
	launchTool(
		targetToolId: string,
		prefillData?: Record<string, unknown>,
		options?: {
			primaryConnectionId?: string | null;
			secondaryConnectionId?: string | null;
		},
	): Promise<unknown>;
	getLaunchContext(): Promise<Record<string, unknown> | null>;
	returnData(data: Record<string, unknown>): Promise<void>;
}

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

/** Safely obtain the invocation API, or undefined when unavailable. */
function getInvocationAPI(): InvocationAPI | undefined {
	if (typeof window === "undefined" || !window.toolboxAPI) return undefined;
	const extended = window.toolboxAPI as typeof window.toolboxAPI & { invocation?: InvocationAPI };
	return extended.invocation;
}

/**
 * Returns true when the PPTB host exposes the T2T invocation API.
 * Use this to conditionally render T2T UI elements.
 */
export function isT2TSupported(): boolean {
	return getInvocationAPI() !== undefined;
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
	});
}
