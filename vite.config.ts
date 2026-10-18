/// <reference types="vitest/config" />
import { readFileSync } from "node:fs";
import { defineConfig, type Plugin } from "vite";
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

// https://vite.dev/config/
export default defineConfig({
	plugins: [react(), toolIcon()],
	optimizeDeps: {
		// Pre-bundle monaco-editor so Vite doesn't try to transform its deep
		// internal imports at runtime (avoids CJS/ESM interop issues).
		include: ["monaco-editor"],
	},
	build: {
		// Disable source maps in production — Monaco worker maps alone add ~17 MB.
		sourcemap: false,
		rollupOptions: {
			output: {
				// Keep Monaco's large language packs in a separate chunk so the
				// main bundle doesn't balloon unnecessarily.
				manualChunks: {
					"monaco-editor": ["monaco-editor"],
				},
			},
		},
	},
	test: {
		// jsdom provides DOMParser, which the FetchXML parser relies on.
		environment: "jsdom",
		include: ["src/**/*.test.ts"],
	},
});
