/**
 * Vite dev-server plugin for the FXS dev harness (`npm run dev:harness`).
 *
 *  - Injects dev/harness/pptbStub.ts into index.html ahead of /src/main.tsx, so the unchanged
 *    app boots against stand-ins for `window.toolboxAPI` / `window.dataverseAPI`.
 *  - Proxies `/__dv/*` → `<org>/api/data/v9.2/*` and attaches `Authorization: Bearer <token>`,
 *    where the token comes from the developer's own Azure CLI session
 *    (`az account get-access-token --resource <org>`) on the *server* side.
 *  - Optionally adds `CallerObjectId: <guid>` to every request, like PPTB does when a user is
 *    impersonated (https://learn.microsoft.com/power-apps/developer/data-platform/webapi/impersonate-another-user-web-api).
 *
 * Safety rails:
 *  - only registered for `vite serve --mode harness` (see vite.config.ts) and `apply: "serve"`
 *  - refuses to start if the dev server is configured to listen beyond loopback
 *  - only answers loopback clients with a loopback Host header, a loopback (or absent) Origin and
 *    the `x-fxs-harness` header, so other web pages can't drive the proxy (CSRF / DNS rebinding)
 *  - the token is cached in memory only, never logged, never written to disk and never sent to
 *    the browser; upstream URLs are pinned to the org's /api/data/v9.2/ path
 */
import { exec, execFile } from "node:child_process";
import { existsSync } from "node:fs";
import type { IncomingMessage, ServerResponse } from "node:http";
import type { Plugin } from "vite";

const API_PATH = "/api/data/v9.2";
const HARNESS_HEADER = "x-fxs-harness";
const STUB_MODULE = "/dev/harness/pptbStub.ts";
const REFRESH_MARGIN_MS = 5 * 60 * 1000;
const LOG_PREFIX = "[fxs-harness]";
/** Default Azure CLI location on Windows; it is often not on PATH. */
const WINDOWS_AZ_DEFAULT = "C:\\Program Files\\Microsoft SDKs\\Azure\\CLI2\\wbin\\az.cmd";

/** Request headers forwarded to Dataverse (plus any `mscrm.*` header). */
const FORWARD_REQUEST_HEADERS = [
	"accept",
	"content-type",
	"prefer",
	"if-match",
	"if-none-match",
	"odata-maxversion",
	"odata-version",
	"consistency",
];

/** Response headers passed back to the browser. */
const FORWARD_RESPONSE_HEADERS = [
	"content-type",
	"odata-entityid",
	"odata-version",
	"preference-applied",
	"location",
	"etag",
	"retry-after",
	"x-ms-service-request-id",
	"req_id",
];

const LOOPBACK_HOSTNAMES = new Set(["localhost", "127.0.0.1", "[::1]", "::1"]);
const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Validates and normalizes DATAVERSE_URL to an origin like https://org.crm.dynamics.com */
export function normalizeOrgUrl(raw: string): string {
	let url: URL;
	try {
		url = new URL(raw.trim());
	} catch {
		throw new Error(`${LOG_PREFIX} DATAVERSE_URL is not a valid URL: ${raw}`);
	}
	if (url.protocol !== "https:" || !/^[a-z0-9-]+(\.[a-z0-9-]+)+$/i.test(url.hostname)) {
		throw new Error(
			`${LOG_PREFIX} DATAVERSE_URL must look like https://<org>.crm.dynamics.com (got ${raw})`,
		);
	}
	return url.origin;
}

/** Validates the optional impersonation id (the user's Microsoft Entra object id). */
export function normalizeCallerObjectId(raw: string | undefined): string | undefined {
	const value = raw?.trim();
	if (!value) return undefined;
	if (!GUID_PATTERN.test(value)) {
		throw new Error(`${LOG_PREFIX} DATAVERSE_CALLER_OBJECT_ID must be a GUID (got ${raw})`);
	}
	return value;
}

function hostnameOf(hostHeader: string | undefined): string {
	if (!hostHeader) return "";
	if (hostHeader.startsWith("[")) return hostHeader.slice(0, hostHeader.indexOf("]") + 1);
	return hostHeader.split(":")[0].toLowerCase();
}

function isLoopbackAddress(address: string | undefined): boolean {
	return address === "127.0.0.1" || address === "::1" || address === "::ffff:127.0.0.1";
}

/** Rejects anything that isn't the harness page itself talking to the dev server. */
export function isTrustedRequest(req: {
	headers: Record<string, string | string[] | undefined>;
	socket?: { remoteAddress?: string };
}): boolean {
	if (req.socket && !isLoopbackAddress(req.socket.remoteAddress)) return false;
	const host = req.headers.host;
	if (typeof host !== "string" || !LOOPBACK_HOSTNAMES.has(hostnameOf(host))) return false;
	if (req.headers[HARNESS_HEADER] !== "1") return false;
	const origin = req.headers.origin;
	if (typeof origin === "string") {
		try {
			if (!LOOPBACK_HOSTNAMES.has(new URL(origin).hostname.toLowerCase())) return false;
		} catch {
			return false;
		}
	}
	return true;
}

/** Builds the upstream URL and guarantees it stays under <org>/api/data/v9.2/. */
export function buildUpstreamUrl(orgOrigin: string, reqUrl: string): URL {
	const target = new URL(`${orgOrigin}${API_PATH}${reqUrl.startsWith("/") ? "" : "/"}${reqUrl}`);
	if (target.origin !== orgOrigin || !target.pathname.startsWith(`${API_PATH}/`)) {
		throw new Error("Path escapes the Dataverse Web API root");
	}
	return target;
}

/** Inserts the stub as the first module script, so it runs before /src/main.tsx. */
export function injectStub(html: string): string {
	const tag = `<script type="module" src="${STUB_MODULE}"></script>`;
	return html.replace(/<script\s+type="module"/, `${tag}\n\t\t<script type="module"`);
}

/** Resolves the az executable: AZURE_CLI_PATH, then the Windows default location, then PATH. */
export function resolveAzCommand(
	env: Record<string, string | undefined> = process.env,
	platform: NodeJS.Platform = process.platform,
	exists: (path: string) => boolean = existsSync,
): string {
	const configured = env.AZURE_CLI_PATH?.trim();
	if (configured) {
		// The path is placed inside double quotes on a cmd.exe command line (see runAz).
		if (/["%!^&|<>\r\n]/.test(configured)) {
			throw new Error(`${LOG_PREFIX} AZURE_CLI_PATH contains characters that aren't allowed`);
		}
		return configured;
	}
	if (platform === "win32" && exists(WINDOWS_AZ_DEFAULT)) return WINDOWS_AZ_DEFAULT;
	return "az";
}

interface CachedToken {
	value: string;
	expiresAt: number;
}

function runAz(args: string[]): Promise<string> {
	const command = resolveAzCommand();
	return new Promise((resolve, reject) => {
		const done = (error: Error | null, stdout: string | Buffer, stderr: string | Buffer) => {
			if (error) {
				// stderr carries az's own error text (never a token). Keep it short.
				const detail = String(stderr).trim().slice(0, 500);
				reject(
					new Error(
						`${LOG_PREFIX} 'az account get-access-token' failed. Run 'az login' ` +
							`(add --tenant <id> if needed) first. ${detail}`,
					),
				);
				return;
			}
			resolve(String(stdout));
		};
		const options = { timeout: 60_000, windowsHide: true, maxBuffer: 1024 * 1024 };
		if (process.platform === "win32") {
			// az is a .cmd shim on Windows, which Node can only start through a shell. Build the
			// command line ourselves: a quoted, validated executable path plus fixed arguments and
			// the validated org origin (no shell metacharacters can reach cmd.exe).
			exec(`"${command}" ${args.join(" ")}`, options, done);
		} else {
			execFile(command, args, options, done);
		}
	});
}

export function createTokenProvider(
	orgOrigin: string,
	run: (args: string[]) => Promise<string> = runAz,
	now: () => number = Date.now,
) {
	let cached: CachedToken | undefined;
	let pending: Promise<string> | undefined;

	async function fetchToken(): Promise<string> {
		// Full JSON (no --query) so we also get the expiry; nothing here is logged.
		const stdout = await run([
			"account",
			"get-access-token",
			"--resource",
			orgOrigin,
			"-o",
			"json",
		]);
		const parsed = JSON.parse(stdout) as {
			accessToken?: string;
			expires_on?: number | string | null;
			expiresOn?: string | null;
		};
		if (!parsed.accessToken) throw new Error(`${LOG_PREFIX} az returned no access token`);
		let expiresAt = Number(parsed.expires_on) * 1000;
		if (!Number.isFinite(expiresAt) || expiresAt <= 0) {
			expiresAt = parsed.expiresOn ? Date.parse(parsed.expiresOn) : NaN;
		}
		if (!Number.isFinite(expiresAt)) expiresAt = now() + 30 * 60 * 1000;
		cached = { value: parsed.accessToken, expiresAt };
		return cached.value;
	}

	return async function getToken(): Promise<string> {
		if (cached && cached.expiresAt - REFRESH_MARGIN_MS > now()) return cached.value;
		pending ??= fetchToken().finally(() => {
			pending = undefined;
		});
		return pending;
	};
}

function readBody(req: IncomingMessage): Promise<Buffer> {
	return new Promise((resolve, reject) => {
		const chunks: Buffer[] = [];
		req.on("data", (c: Buffer) => chunks.push(c));
		req.on("end", () => resolve(Buffer.concat(chunks)));
		req.on("error", reject);
	});
}

function sendJsonError(res: ServerResponse, status: number, code: string, message: string) {
	res.statusCode = status;
	res.setHeader("content-type", "application/json; charset=utf-8");
	res.end(JSON.stringify({ error: { code, message } }));
}

export interface DataverseProxyOptions {
	/** e.g. https://contoso.crm.dynamics.com (from DATAVERSE_URL) */
	orgUrl: string;
	/** Optional Microsoft Entra object id to impersonate (from DATAVERSE_CALLER_OBJECT_ID). */
	callerObjectId?: string;
}

export function dataverseHarness(options: DataverseProxyOptions): Plugin {
	const orgOrigin = normalizeOrgUrl(options.orgUrl);
	const callerObjectId = normalizeCallerObjectId(options.callerObjectId);
	const getToken = createTokenProvider(orgOrigin);

	return {
		name: "fxs-dev-harness",
		apply: "serve",
		transformIndexHtml: {
			order: "pre",
			handler: (html) => injectStub(html),
		},
		configureServer(server) {
			const host = server.config.server.host;
			if (host === true || (typeof host === "string" && !LOOPBACK_HOSTNAMES.has(host))) {
				throw new Error(
					`${LOG_PREFIX} refusing to proxy Dataverse while the dev server listens on '${String(host)}'. ` +
						"Run the harness on localhost only.",
				);
			}

			server.config.logger.info(
				`${LOG_PREFIX} proxying /__dv/* → ${orgOrigin}${API_PATH}/*` +
					(callerObjectId ? " (impersonating via CallerObjectId)" : ""),
			);

			server.middlewares.use("/__dv-harness/info", (req, res) => {
				if (!isTrustedRequest(req)) {
					sendJsonError(res, 403, "Forbidden", "Untrusted request");
					return;
				}
				res.setHeader("content-type", "application/json; charset=utf-8");
				res.end(JSON.stringify({ orgUrl: orgOrigin, impersonating: !!callerObjectId }));
			});

			server.middlewares.use("/__dv", (req, res) => {
				void (async () => {
					if (!isTrustedRequest(req)) {
						sendJsonError(res, 403, "Forbidden", "Untrusted request");
						return;
					}

					let target: URL;
					try {
						target = buildUpstreamUrl(orgOrigin, req.url ?? "/");
					} catch (error) {
						sendJsonError(res, 400, "BadRequest", (error as Error).message);
						return;
					}

					const headers: Record<string, string> = {};
					for (const [name, value] of Object.entries(req.headers)) {
						const lower = name.toLowerCase();
						if (typeof value !== "string") continue;
						if (FORWARD_REQUEST_HEADERS.includes(lower) || lower.startsWith("mscrm.")) {
							headers[lower] = value;
						}
					}
					if (callerObjectId) headers.callerobjectid = callerObjectId;

					const method = (req.method ?? "GET").toUpperCase();
					const body =
						method === "GET" || method === "HEAD" ? undefined : await readBody(req);

					let token: string;
					try {
						token = await getToken();
					} catch (error) {
						server.config.logger.error((error as Error).message);
						sendJsonError(res, 502, "HarnessAuthError", (error as Error).message);
						return;
					}
					headers.authorization = `Bearer ${token}`;

					const started = Date.now();
					const upstream = await fetch(target, {
						method,
						headers,
						body: body && body.length > 0 ? new Uint8Array(body) : undefined,
						redirect: "manual",
					});
					const payload = Buffer.from(await upstream.arrayBuffer());

					res.statusCode = upstream.status;
					for (const name of FORWARD_RESPONSE_HEADERS) {
						const value = upstream.headers.get(name);
						if (value !== null) res.setHeader(name, value);
					}
					res.end(payload);

					// Method, path and status only — never headers, bodies, query strings or the token.
					server.config.logger.info(
						`${LOG_PREFIX} ${method} ${target.pathname.slice(API_PATH.length)} → ${upstream.status} (${Date.now() - started} ms)`,
					);
				})().catch((error: unknown) => {
					server.config.logger.error(
						`${LOG_PREFIX} proxy error: ${(error as Error).message}`,
					);
					if (!res.headersSent) {
						sendJsonError(res, 502, "HarnessProxyError", (error as Error).message);
					} else {
						res.end();
					}
				});
			});
		},
	};
}
