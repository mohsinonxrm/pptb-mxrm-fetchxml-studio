/**
 * Dev harness: installs stand-ins for PPTB's `window.toolboxAPI` and `window.dataverseAPI` so
 * FetchXML Studio's real UI runs in a normal browser tab against a live Dataverse org.
 *
 * Injected into index.html, ahead of /src/main.tsx, by the harness Vite plugin
 * (vite-plugin-dataverse-proxy.ts) — only under `npm run dev:harness`. Never part of the build
 * or the npm package.
 *
 * Only the host members FXS uses are implemented; anything else is undefined, so FXS's own
 * optional-chaining fallbacks apply. Dataverse calls go through the dev-server proxy (`/__dv/*`),
 * which adds the bearer token server-side.
 *
 * Query-string options (handy for screenshots and automation):
 *   ?theme=dark|light        force the theme returned by utils.getCurrentTheme()
 *   ?fetchXml=<xml>          simulate a Tool-to-Tool launch with this FetchXML prefill
 *   ?prefill=<json>          simulate a T2T launch with an arbitrary prefill object
 *                            (e.g. {"viewRef":{"id":"…","entityLogicalName":"savedquery"}})
 */
import { createDataverseStub, HARNESS_HEADER } from "./dataverseStub";

const LOG_PREFIX = "[fxs-harness]";
const TOOL_ID = "fxs-dev-harness";
const SETTINGS_PREFIX = "fxs-harness:settings:";

interface HarnessInfo {
	orgUrl: string;
	impersonating: boolean;
}

let infoPromise: Promise<HarnessInfo> | undefined;

function getHarnessInfo(): Promise<HarnessInfo> {
	infoPromise ??= fetch("/__dv-harness/info", { headers: { [HARNESS_HEADER]: "1" } }).then(
		async (res) => {
			if (!res.ok) {
				throw new Error(
					`Dev harness proxy is not running (HTTP ${res.status}). ` +
						"Start it with DATAVERSE_URL=https://<org>.crm.dynamics.com npm run dev:harness",
				);
			}
			return (await res.json()) as HarnessInfo;
		},
	);
	return infoPromise;
}

const params = new URLSearchParams(window.location.search);

function currentTheme(): "light" | "dark" {
	const forced = params.get("theme");
	if (forced === "dark" || forced === "light") return forced;
	return window.matchMedia?.("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

/** Simulated T2T launch context from the query string; null for a standalone launch. */
function launchContextFromQuery(): Record<string, unknown> | null {
	const fetchXml = params.get("fetchXml");
	if (fetchXml) return { fetchXml };
	const prefill = params.get("prefill");
	if (!prefill) return null;
	try {
		const parsed: unknown = JSON.parse(prefill);
		return parsed && typeof parsed === "object" && !Array.isArray(parsed)
			? (parsed as Record<string, unknown>)
			: null;
	} catch {
		console.warn(LOG_PREFIX, "?prefill is not valid JSON; ignoring it.");
		return null;
	}
}

// PPTB persists tool settings per tool; localStorage is the closest browser equivalent.
function readSetting(key: string): unknown {
	const raw = window.localStorage.getItem(SETTINGS_PREFIX + key);
	if (raw === null) return undefined;
	try {
		return JSON.parse(raw);
	} catch {
		return undefined;
	}
}

function writeSetting(key: string, value: unknown): void {
	window.localStorage.setItem(SETTINGS_PREFIX + key, JSON.stringify(value));
}

/** Browser download in place of the host's native save dialog. Returns the file name. */
function download(fileName: string, content: unknown): string {
	const isBlobPart =
		typeof content === "string" ||
		content instanceof Blob ||
		content instanceof ArrayBuffer ||
		ArrayBuffer.isView(content);
	const part = isBlobPart ? (content as BlobPart) : JSON.stringify(content);
	const url = URL.createObjectURL(new Blob([part]));
	const link = document.createElement("a");
	link.href = url;
	link.download = fileName;
	document.body.appendChild(link);
	link.click();
	link.remove();
	URL.revokeObjectURL(url);
	return fileName;
}

const dataverseAPI = createDataverseStub((input, init) => window.fetch(input, init));

const toolboxAPI = {
	connections: {
		getActiveConnection: async (): Promise<ToolBoxAPI.DataverseConnection> => {
			const { orgUrl, impersonating } = await getHarnessInfo();
			return {
				id: TOOL_ID,
				name: `${new URL(orgUrl).hostname} (dev harness${impersonating ? ", impersonating" : ""})`,
				url: orgUrl,
				environment: "Dev",
			};
		},
		getSecondaryConnection: async (): Promise<null> => null,
	},
	utils: {
		getCurrentTheme: async (): Promise<"light" | "dark"> => currentTheme(),
		copyToClipboard: async (text: string): Promise<void> => {
			await navigator.clipboard.writeText(text);
		},
		showNotification: async (options: ToolBoxAPI.NotificationOptions): Promise<void> => {
			console.info(
				LOG_PREFIX,
				`[${options.type ?? "info"}] ${options.title}: ${options.body}`,
			);
		},
	},
	settings: {
		get: async (key: string): Promise<unknown> => readSetting(key),
		set: async (key: string, value: unknown): Promise<void> => writeSetting(key, value),
		getAll: async (): Promise<Record<string, unknown>> => {
			const all: Record<string, unknown> = {};
			for (let i = 0; i < window.localStorage.length; i++) {
				const storageKey = window.localStorage.key(i);
				if (storageKey?.startsWith(SETTINGS_PREFIX)) {
					const key = storageKey.slice(SETTINGS_PREFIX.length);
					all[key] = readSetting(key);
				}
			}
			return all;
		},
		setAll: async (settings: Record<string, unknown>): Promise<void> => {
			for (const [key, value] of Object.entries(settings)) writeSetting(key, value);
		},
	},
	fileSystem: {
		saveFile: async (defaultPath: string, content: unknown): Promise<string | null> =>
			download(defaultPath, content),
	},
	events: {
		// The harness emits no host events (no connection switching, settings changes, …).
		on: (): void => undefined,
		off: (): void => undefined,
	},
	getToolContext: async (): Promise<ToolBoxAPI.ToolContext> => {
		const { orgUrl } = await getHarnessInfo();
		return { toolId: TOOL_ID, connectionUrl: orgUrl, connectionId: TOOL_ID };
	},
	invocation: {
		getLaunchContext: async (): Promise<Record<string, unknown> | null> =>
			launchContextFromQuery(),
		returnData: async (data: Record<string, unknown>): Promise<void> => {
			console.info(LOG_PREFIX, "invocation.returnData:", data);
		},
		launchTool: async (targetToolId: string): Promise<unknown> => {
			throw new Error(
				`Tool-to-Tool launch (${targetToolId}) is not available in the dev harness`,
			);
		},
		// No other tools are installed in the harness.
		findToolsByCapability: async (): Promise<unknown[]> => [],
		getKnownCapabilityTags: async (): Promise<
			Array<{ tag: string; description: string }>
		> => [],
	},
};

// The stubs implement only what FXS uses, so cast to the full host API types.
window.toolboxAPI = toolboxAPI as unknown as ToolBoxAPI.API;
window.dataverseAPI = dataverseAPI as unknown as DataverseAPI.API;

console.info(LOG_PREFIX, "PPTB stub installed (toolboxAPI + dataverseAPI → /__dv proxy).");
