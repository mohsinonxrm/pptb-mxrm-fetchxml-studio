import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// https://vite.dev/config/
export default defineConfig({
	plugins: [react()],
	optimizeDeps: {
		// Include monaco-editor in pre-bundling to avoid issues
		include: ["monaco-editor"],
	},
	build: {
		rollupOptions: {
			// Ensure workers are properly handled in production build
			output: {
				manualChunks: {
					"monaco-editor": ["monaco-editor"],
				},
			},
		},
	},
});
