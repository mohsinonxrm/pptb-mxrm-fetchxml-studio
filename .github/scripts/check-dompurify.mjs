// Fails when the built bundle carries a DOMPurify other than the installed one.
// Monaco ships its own copy; vite.config.ts (monacoDompurify) replaces it with the npm package,
// whose version package.json's `overrides` sets. Same check as Command Bar Studio 0.6.2.
// usage: node .github/scripts/check-dompurify.mjs [distDir]
import { readFileSync, readdirSync, statSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

const dist = process.argv[2] ?? "dist";
// The version Monaco resolves: its own dependency, as the override set it. Monaco's folder is
// found on disk: 0.57's `exports` map hides monaco-editor/package.json.
const fromMonaco = createRequire(path.join(process.cwd(), "node_modules", "monaco-editor", "package.json"));
let dir = path.dirname(fromMonaco.resolve("dompurify"));
while (!statSync(path.join(dir, "package.json"), { throwIfNoEntry: false })) dir = path.dirname(dir);
const installed = JSON.parse(readFileSync(path.join(dir, "package.json"), "utf8")).version;

const files = [];
const walk = (d) => {
	for (const name of readdirSync(d)) {
		const p = path.join(d, name);
		if (statSync(p).isDirectory()) walk(p);
		else if (p.endsWith(".js")) files.push(p);
	}
};
walk(dist);

// DOMPurify sets `version` then `removed` on its instance; minified: e.version="3.4.17",e.removed=[]
const found = new Map();
for (const f of files) {
	for (const m of readFileSync(f, "utf8").matchAll(/\.version\s*=\s*["'](\d+\.\d+\.\d+)["']\s*,\s*\w+\.removed\s*=\s*\[\]/g)) {
		const key = path.relative(dist, f);
		found.set(key, [...(found.get(key) ?? []), m[1]]);
	}
}
if (found.size === 0) {
	console.error(`No DOMPurify found in ${dist}: the check can't verify the bundle.`);
	process.exit(1);
}
const wrong = [...found].filter(([, versions]) => versions.some((v) => v !== installed));
for (const [file, versions] of found) console.log(`${file}: DOMPurify ${versions.join(", ")}`);
if (wrong.length > 0) {
	console.error(
		`The bundle has DOMPurify ${[...new Set(wrong.flatMap(([, v]) => v))].join(", ")}, not the installed ${installed}.`,
	);
	process.exit(1);
}
console.log(`OK: the bundle's DOMPurify is the installed ${installed}.`);
