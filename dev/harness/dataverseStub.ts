/**
 * Browser-side stand-in for PPTB's `window.dataverseAPI`, used only by the dev harness
 * (`npm run dev:harness`). Never shipped: `dev/` is not a build input and not in the package.
 *
 * Request shaping mirrors the PPTB desktop host (PowerPlatformToolBox/desktop-app v1.2.7,
 * src/main/managers/dataverseManager.ts), so host behaviour (and host bugs) reproduce here:
 *  - entity-set names come from the host's naive pluralization (getEntitySetName), used by
 *    create / update / delete / fetchXmlQuery and bound execute() operations
 *  - execute() builds bound/unbound URLs and function parameter aliases the same way
 *    (formatFunctionParameter, including unquoted `Microsoft.Dynamics.CRM.X'Y'` enum literals)
 *  - every request sends `Prefer: return=representation`; queryData() and fetchXmlQuery() also
 *    ask for `odata.include-annotations="*"`. A caller-supplied Prefer header replaces the extras.
 *  - non-2xx responses reject with `"<code>: <message>"` (or `"HTTP <status>: <body>"`)
 *
 * Requests go to the Vite dev-server proxy at `/__dv/*`, which forwards them to
 * `<org>/api/data/v9.2/*` and adds the bearer token (and CallerObjectId, when configured)
 * server-side. The token never reaches the browser.
 *
 * Only the members FetchXML Studio calls are implemented (see pptbStub.ts).
 */

/** Header the proxy requires on every request (forces a CORS preflight, so blocks CSRF). */
export const HARNESS_HEADER = "x-fxs-harness";

export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

export interface ExecuteRequestLike {
	entityName?: string;
	entityId?: string;
	operationName: string;
	operationType: "action" | "function";
	parameters?: Record<string, unknown>;
}

type HeaderMap = Record<string, string>;
type ConnectionTarget = "primary" | "secondary";

const IRREGULAR_PLURALS: Record<string, string> = {
	opportunity: "opportunities",
	territory: "territories",
	currency: "currencies",
	businessunit: "businessunits",
	systemuser: "systemusers",
	usersettingscollection: "usersettingscollection",
	principalobjectaccess: "principalobjectaccessset",
	webresource: "webresourceset",
	relationshipdefinition: "RelationshipDefinitions",
};

/** Same pluralization rules as PPTB's DataverseManager.getEntitySetName. */
export function getEntitySetName(entityLogicalName: string): string {
	const lowerName = entityLogicalName.toLowerCase();
	const irregular = IRREGULAR_PLURALS[lowerName];
	if (irregular) return irregular;
	if (
		lowerName.endsWith("y") &&
		lowerName.length > 1 &&
		!"aeiou".includes(lowerName[lowerName.length - 2])
	) {
		return lowerName.slice(0, -1) + "ies";
	}
	if (
		lowerName.endsWith("s") ||
		lowerName.endsWith("x") ||
		lowerName.endsWith("z") ||
		lowerName.endsWith("ch") ||
		lowerName.endsWith("sh")
	) {
		return lowerName + "es";
	}
	return lowerName + "s";
}

/** Same encoding as PPTB's DataverseManager.formatFunctionParameter. */
export function formatFunctionParameter(value: unknown): string {
	if (value === null || value === undefined) return "null";

	if (typeof value === "object" && "entityLogicalName" in value && "id" in value) {
		const ref = value as { entityLogicalName: unknown; id: unknown };
		if (typeof ref.entityLogicalName !== "string" || typeof ref.id !== "string") {
			throw new Error("EntityReference must have string entityLogicalName and id properties");
		}
		const odataRef = { "@odata.id": `${getEntitySetName(ref.entityLogicalName)}(${ref.id})` };
		return encodeURIComponent(JSON.stringify(odataRef));
	}

	if (typeof value === "object" && "@odata.id" in value) {
		return encodeURIComponent(JSON.stringify(value));
	}

	if (typeof value === "boolean") return value ? "true" : "false";
	if (typeof value === "number") return value.toString();
	if (value instanceof Date) return value.toISOString();

	if (typeof value === "string") {
		// Enum literal, e.g. Microsoft.Dynamics.CRM.EntityFilters'Entity' — not quoted.
		if (/^Microsoft\.Dynamics\.CRM\.\w+'.+'$/.test(value)) return encodeURIComponent(value);
		if (value.startsWith("{'@odata.id':") || value.startsWith('{"@odata.id":')) {
			return encodeURIComponent(value);
		}
		return encodeURIComponent(`'${value.replace(/'/g, "''")}'`);
	}

	if (typeof value === "object") return encodeURIComponent(JSON.stringify(value));
	return encodeURIComponent(String(value));
}

/**
 * Builds the Web API path (relative to `/api/data/v9.2/`), HTTP method and body for an
 * execute() request — same rules as PPTB's DataverseManager.execute.
 */
export function buildExecuteRequest(request: ExecuteRequestLike): {
	path: string;
	method: "GET" | "POST";
	body?: Record<string, unknown>;
} {
	let path = "";
	if (request.entityName) {
		const entitySetName = getEntitySetName(request.entityName);
		path += request.entityId
			? `${entitySetName}(${request.entityId})/Microsoft.Dynamics.CRM.${request.operationName}`
			: `${entitySetName}/Microsoft.Dynamics.CRM.${request.operationName}`;
	} else {
		path += request.operationName;
	}

	const isFunction = request.operationType === "function";
	if (isFunction && request.parameters) {
		const names = Object.keys(request.parameters);
		if (names.length > 0) {
			path += `(${names.map((name, i) => `${name}=@p${i}`).join(",")})`;
			path += `?${names
				.map((name, i) => `@p${i}=${formatFunctionParameter(request.parameters![name])}`)
				.join("&")}`;
		} else {
			path += "()";
		}
	}

	return {
		path,
		method: isFunction ? "GET" : "POST",
		body: request.operationType === "action" ? request.parameters : undefined,
	};
}

function extractIdFromUrl(url: string): string {
	const match = url.match(/\(([a-f0-9-]+)\)/i);
	return match ? match[1] : url;
}

interface HttpResult {
	status: number;
	data: Record<string, unknown>;
	headers: Headers;
}

/**
 * Creates the dataverseAPI stand-in. `fetchImpl` and `basePath` are injectable for tests; at
 * runtime they are `window.fetch` and the `/__dv` proxy.
 *
 * The optional trailing `additionalHeaders` argument mirrors the host's header pass-through
 * (PPTB merges caller headers and lets a custom Prefer replace the default extras).
 */
export function createDataverseStub(fetchImpl: FetchLike, basePath = "/__dv") {
	async function request(
		path: string,
		method: string,
		body?: unknown,
		preferOptions?: string[],
		additionalHeaders?: HeaderMap,
	): Promise<HttpResult> {
		const custom: HeaderMap = { ...(additionalHeaders ?? {}) };
		const prefer = ["return=representation"];
		const customPreferKey = Object.keys(custom).find((k) => k.toLowerCase() === "prefer");
		if (customPreferKey && custom[customPreferKey]) {
			prefer.push(custom[customPreferKey]);
			delete custom[customPreferKey];
		} else if (preferOptions?.length) {
			prefer.push(...preferOptions);
		}

		const response = await fetchImpl(`${basePath}/${path.replace(/^\//, "")}`, {
			method,
			headers: {
				Accept: "application/json",
				"OData-MaxVersion": "4.0",
				"OData-Version": "4.0",
				"Content-Type": "application/json; charset=utf-8",
				Prefer: prefer.join(","),
				[HARNESS_HEADER]: "1",
				...custom,
			},
			body:
				body === undefined
					? undefined
					: typeof body === "string"
						? body
						: JSON.stringify(body),
		});

		const text = await response.text();
		if (response.ok) {
			let data: Record<string, unknown> = {};
			if (text.trim()) {
				try {
					data = JSON.parse(text) as Record<string, unknown>;
				} catch {
					data = {};
				}
			}
			return { status: response.status, data, headers: response.headers };
		}

		let message = `HTTP ${response.status}`;
		try {
			const err = JSON.parse(text) as { error?: { code?: string; message?: string } };
			if (err.error) message = `${err.error.code}: ${err.error.message}`;
		} catch {
			if (text.trim()) message += `: ${text.trim()}`;
		}
		throw new Error(message);
	}

	const annotations = ['odata.include-annotations="*"'];

	return {
		queryData: async (
			odataQuery: string,
			_connectionTarget?: ConnectionTarget,
			additionalHeaders?: HeaderMap,
		): Promise<{ value: Record<string, unknown>[] }> => {
			if (!odataQuery || !odataQuery.trim()) {
				throw new Error("odataQuery parameter cannot be empty");
			}
			const query = odataQuery.trim();
			const clean = query.startsWith("?") ? query.substring(1) : query;
			const result = await request(clean, "GET", undefined, annotations, additionalHeaders);
			return result.data as { value: Record<string, unknown>[] };
		},

		fetchXmlQuery: async (
			fetchXml: string,
			_connectionTarget?: ConnectionTarget,
			additionalHeaders?: HeaderMap,
		): Promise<{ value: Record<string, unknown>[] }> => {
			const entityMatch = fetchXml.match(/<entity\s+name=["']([^"']+)["']/i);
			if (!entityMatch) throw new Error("Invalid FetchXML: Could not determine entity name");
			const path = `${getEntitySetName(entityMatch[1])}?fetchXml=${encodeURIComponent(fetchXml)}`;
			const result = await request(path, "GET", undefined, annotations, additionalHeaders);
			return result.data as { value: Record<string, unknown>[] };
		},

		getEntityMetadata: async (
			entityLogicalNameOrId: string,
			searchByLogicalName: boolean,
			entityProperties?: string[],
			_connectionTarget?: ConnectionTarget,
			additionalHeaders?: HeaderMap,
		): Promise<Record<string, unknown>> => {
			if (!entityLogicalNameOrId || !entityLogicalNameOrId.trim()) {
				throw new Error("entityLogicalName parameter cannot be empty");
			}
			const encoded = encodeURIComponent(entityLogicalNameOrId);
			let path = `EntityDefinitions(${searchByLogicalName ? `LogicalName='${encoded}'` : encoded})`;
			if (entityProperties && entityProperties.length > 0) {
				path += `?$select=${entityProperties.map((c) => encodeURIComponent(c)).join(",")}`;
			}
			return (await request(path, "GET", undefined, undefined, additionalHeaders)).data;
		},

		execute: async (
			req: ExecuteRequestLike,
			_connectionTarget?: ConnectionTarget,
			additionalHeaders?: HeaderMap,
		): Promise<Record<string, unknown>> => {
			const { path, method, body } = buildExecuteRequest(req);
			return (await request(path, method, body, undefined, additionalHeaders)).data;
		},

		create: async (
			entityLogicalName: string,
			record: Record<string, unknown>,
			_connectionTarget?: ConnectionTarget,
			additionalHeaders?: HeaderMap,
		): Promise<{ id: string; [key: string]: unknown }> => {
			const result = await request(
				getEntitySetName(entityLogicalName),
				"POST",
				record,
				undefined,
				additionalHeaders,
			);
			const entityId =
				result.headers.get("odata-entityid") ??
				(result.data[`${entityLogicalName}id`] as string | undefined);
			return { id: entityId ? extractIdFromUrl(entityId) : "", ...result.data };
		},

		update: async (
			entityLogicalName: string,
			id: string,
			record: Record<string, unknown>,
			_connectionTarget?: ConnectionTarget,
			additionalHeaders?: HeaderMap,
		): Promise<void> => {
			await request(
				`${getEntitySetName(entityLogicalName)}(${id})`,
				"PATCH",
				record,
				undefined,
				additionalHeaders,
			);
		},

		delete: async (
			entityLogicalName: string,
			id: string,
			_connectionTarget?: ConnectionTarget,
			additionalHeaders?: HeaderMap,
		): Promise<void> => {
			await request(
				`${getEntitySetName(entityLogicalName)}(${id})`,
				"DELETE",
				undefined,
				undefined,
				additionalHeaders,
			);
		},

		getEntitySetName: async (entityLogicalName: string): Promise<string> =>
			getEntitySetName(entityLogicalName),
	};
}

export type DataverseStub = ReturnType<typeof createDataverseStub>;
