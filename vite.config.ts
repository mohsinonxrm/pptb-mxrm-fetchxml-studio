/// <reference types="vitest/config" />
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { defineConfig, loadEnv, type Plugin } from "vite";
import react from "@vitejs/plugin-react";

/**
 * Bundle the tool icon at dist/icons/icon.svg, where package.json's top-level `icon` points.
 * PPTB resolves `icon` relative to the dist root (Verified requirement: an SVG bundled under
 * dist). Prerelease versions (e.g. 1.4.1-beta.7) get the gold Insider variant; releases get the
 * regular icon. https://docs.powerplatformtoolbox.com/tool-development/publishing
 */
function toolIcon(): Plugin {
	return {
		name: "fxs-tool-icon",
		apply: "build",
		generateBundle() {
			const { version } = JSON.parse(readFileSync("package.json", "utf8")) as { version: string };
			const isPrerelease = version.includes("-");
			const source = isPrerelease ? "icons/icon-insider.svg" : "icons/icon.svg";
			this.emitFile({ type: "asset", fileName: "icons/icon.svg", source: readFileSync(source, "utf8") });
		},
	};
}

/**
 * Monaco ships its own copy of DOMPurify (monaco-editor/esm/vs/base/browser/dompurify/dompurify.js),
 * so npm `overrides` alone never reach the bundle: 1.4.4 shipped 3.1.7 and 1.5.0-beta.15 shipped
 * 3.4.15. This resolves Monaco's import to the npm `dompurify` package instead (same default
 * export, the DOMPurify instance), so security fixes come from npm, at the version package.json's
 * override sets. .github/scripts/check-dompurify.mjs checks the built bundle in CI and before
 * publishing. Same approach as Command Bar Studio 0.6.2.
 */
function monacoDompurify(): Plugin {
	return {
		name: "fxs-monaco-dompurify",
		enforce: "pre",
		async resolveId(source, importer, options) {
			if (!importer || !/[\\/]monaco-editor[\\/]esm[\\/]vs[\\/]base[\\/]browser[\\/]/.test(importer)) return null;
			if (!/(^|[\\/])dompurify[\\/]dompurify\.js$/.test(source)) return null;
			// Resolved from Monaco's folder: its own (overridden) dompurify dependency
			return this.resolve("dompurify", importer, { ...options, skipSelf: true });
		},
	};
}

/**
 *  * Dev harness only (`npm run dev:harness` = `vite --mode harness`): runs the real UI in a browser
 * against a live org through a stubbed PPTB host and an Azure CLI–authenticated proxy (#78).
 * Never loaded for `vite build`, `vite` (plain dev) or Vitest. See dev/README.md.
 */
async function devHarness(command: string, mode: string): Promise<Plugin[]> {
	if (command !== "serve" || mode !== "harness") return [];
	// Not VITE_-prefixed, so these stay server-side and never reach the client bundle.
	// .env.local next to this file, whichever directory vite was started from
	const env = loadEnv(mode, fileURLToPath(new URL(".", import.meta.url)), "");
	if (!env.DATAVERSE_URL) {
		throw new Error(
			"[fxs-harness] Set DATAVERSE_URL=https://<org>.crm.dynamics.com (environment variable or " +
				".env.local) before running `npm run dev:harness`. See dev/README.md.",
		);
	}
	const { dataverseHarness } = await import("./dev/harness/vite-plugin-dataverse-proxy");
	return [
		dataverseHarness({
			orgUrl: env.DATAVERSE_URL,
			callerObjectId: env.DATAVERSE_CALLER_OBJECT_ID,
		}),
	];
}

// https://vite.dev/config/
export default defineConfig(async ({ command, mode }) => ({
	plugins: [react(), toolIcon(), monacoDompurify(), ...(await devHarness(command, mode))],
	optimizeDeps: {
		// Pre-bundle monaco-editor so Vite doesn't try to transform its deep
		// internal imports at runtime (avoids CJS/ESM interop issues).
		// These are the entry points main.tsx imports (#145).
		include: [
			"monaco-editor/editor",
			"monaco-editor/features/register.all",
			"monaco-editor/languages/definitions/xml/register",
			"monaco-editor/languages/definitions/csharp/register",
			"monaco-editor/languages/definitions/javascript/register",
			"monaco-editor/languages/definitions/shell/register",
			"monaco-editor/languages/definitions/sql/register",
		],
	},
	build: {
		// Disable source maps in production — Monaco worker maps alone add ~17 MB.
		sourcemap: false,
		rollupOptions: {
			output: {
				// Keep Monaco's large language packs in a separate chunk so the
				// main bundle doesn't balloon unnecessarily. Match by path: the object form
				// (`{ "monaco-editor": ["monaco-editor"] }`) adds the package's root entry to the
				// build, which pulls in every language and worker that main.tsx leaves out (#145).
				manualChunks(id: string) {
					if (id.replace(/\\/g, "/").includes("/node_modules/monaco-editor/")) {
						return "monaco-editor";
					}
				},
			},
		},
	},
	test: {
		// jsdom provides DOMParser, which the FetchXML parser relies on.
		environment: "jsdom",
		include: ["src/**/*.test.ts", "dev/**/*.test.ts"],
	},
}));
