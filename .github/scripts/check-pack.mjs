#!/usr/bin/env node
// Validates `npm pack --dry-run --json` output: required files are present,
// dev-only/analysis files are absent, and the package stays under a size budget.
//
// Usage: npm pack --dry-run --json --silent > pack.json && node .github/scripts/check-pack.mjs pack.json
import { readFileSync, appendFileSync } from "node:fs";

const MAX_UNPACKED_BYTES = 10 * 1024 * 1024; // current: ~6.6 MB
const MAX_PACKED_BYTES = 3 * 1024 * 1024; // current: ~1.8 MB

const REQUIRED = [
	"package.json",
	"index.html",
	"dist/index.html",
	"dist/icons/icon.svg",
	"pptb.config.json",
	"README.md",
	"CHANGELOG.md",
	"LICENSE",
	"SECURITY.md",
];
const FORBIDDEN = [/^dev\//, /^src\//, /^coverage\//, /\.map$/, /^\.github\//];

const raw = readFileSync(process.argv[2] ?? 0, "utf8");
// npm <= 11 prints an array of packages, npm 12 an object keyed by package name.
// Tolerate stray lifecycle-script output around the JSON.
const start = raw.search(/[[{]/);
const end = Math.max(raw.lastIndexOf("]"), raw.lastIndexOf("}"));
const parsed = JSON.parse(raw.slice(start, end + 1));
const pkg = Array.isArray(parsed) ? parsed[0] : Object.values(parsed)[0];
const files = pkg.files.map((f) => f.path.replace(/\\/g, "/"));

const errors = [];
for (const req of REQUIRED) {
	if (!files.includes(req)) errors.push(`missing required file: ${req}`);
}
for (const file of files) {
	if (FORBIDDEN.some((re) => re.test(file))) errors.push(`forbidden file in package: ${file}`);
}
if (pkg.unpackedSize > MAX_UNPACKED_BYTES) {
	errors.push(`unpacked size ${pkg.unpackedSize} B exceeds budget ${MAX_UNPACKED_BYTES} B`);
}
if (pkg.size > MAX_PACKED_BYTES) {
	errors.push(`packed size ${pkg.size} B exceeds budget ${MAX_PACKED_BYTES} B`);
}

const mb = (n) => `${(n / 1024 / 1024).toFixed(2)} MB`;
const summary = [
	`### 📦 ${pkg.name}@${pkg.version}`,
	"",
	`| files | packed | unpacked |`,
	`| --- | --- | --- |`,
	`| ${pkg.entryCount} | ${mb(pkg.size)} | ${mb(pkg.unpackedSize)} |`,
	"",
	errors.length
		? `❌ ${errors.length} problem(s):\n${errors.map((e) => `- ${e}`).join("\n")}`
		: "✅ package contents OK",
	"",
].join("\n");

console.log(summary);
if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, summary);
if (errors.length) process.exit(1);
