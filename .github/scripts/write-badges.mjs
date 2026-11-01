#!/usr/bin/env node
// Writes assets/badges/*.svg from package.json (`npm run badges`). Run it after changing the
// version or a badge's dependency; check-release.mjs fails when the files are out of date.
import { mkdirSync, writeFileSync, readFileSync } from "node:fs";
import { badges } from "./badges.mjs";

const pkg = JSON.parse(readFileSync("package.json", "utf8"));
mkdirSync("assets/badges", { recursive: true });
for (const { file, svg } of badges(pkg)) {
	writeFileSync(`assets/badges/${file}`, svg);
	console.log(`assets/badges/${file}`);
}
