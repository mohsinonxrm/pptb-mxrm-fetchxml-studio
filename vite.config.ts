import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// https://vite.dev/config/
export default defineConfig({
	plugins: [react()],
	optimizeDeps: {
		// Pre-bundle monaco-editor so Vite doesn't try to transform its deep
		// internal imports at runtime (avoids CJS/ESM interop issues).
		include: ["monaco-editor"],
	},
	build: {
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
});
