#!/usr/bin/env node
// Writes assets/badges/*.svg from package.json (`npm run badges`), points the README's badge lines
// at the version's tag, and sets package.json's `configurations.readmeUrl` for the version (#379):
// `npm run badges -- --support` for a support release (automatic on its support/X.Y branch).
// Run it after changing the version or a badge's dependency; check-release.mjs fails when any of
// them is out of date.
import { execSync } from "node:child_process";
import { mkdirSync, writeFileSync, readFileSync } from "node:fs";
import {
	badgeMarkdown,
	badges,
	missingBadgeLines,
	readmeUrlFor,
	repositoryOf,
	rewriteBadgeLines,
	rewriteReadmeUrl,
	supportBranchOf,
} from "./badges.mjs";

const pkgText = readFileSync("package.json", "utf8");
const pkg = JSON.parse(pkgText);
mkdirSync("assets/badges", { recursive: true });
for (const { file, svg } of badges(pkg)) {
	writeFileSync(`assets/badges/${file}`, svg);
	console.log(`assets/badges/${file}`);
}

const readme = readFileSync("README.md", "utf8");
const nextReadme = rewriteBadgeLines(readme, {
	repository: repositoryOf(pkg),
	files: badges(pkg).map((b) => b.file),
	lines: badgeMarkdown(pkg),
});
if (nextReadme !== readme) writeFileSync("README.md", nextReadme);
console.log("README.md badge lines");
// A new badge has no line to rewrite: say so, rather than let the release check find it later
for (const line of missingBadgeLines(nextReadme, badgeMarkdown(pkg)))
	console.warn(`README.md has no line for this badge; add it with the others: ${line}`);

// A support release reads its README from support/X.Y: with --support, or when run on that branch
const currentBranch = () => {
	try {
		return execSync("git rev-parse --abbrev-ref HEAD", { encoding: "utf8" }).trim();
	} catch {
		return undefined;
	}
};
const support = process.argv.includes("--support") || currentBranch() === supportBranchOf(pkg);
const readmeUrl = readmeUrlFor(pkg, { support });
const nextPkg = rewriteReadmeUrl(pkgText, readmeUrl);
if (nextPkg !== pkgText) writeFileSync("package.json", nextPkg);
if (JSON.parse(nextPkg).configurations?.readmeUrl === readmeUrl)
	console.log(`package.json readmeUrl → ${readmeUrl}`);
else console.warn(`package.json has no configurations.readmeUrl; add it: ${readmeUrl}`);
