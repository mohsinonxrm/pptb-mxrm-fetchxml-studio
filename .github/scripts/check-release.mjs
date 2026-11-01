#!/usr/bin/env node
// Fails when the release metadata doesn't match package.json's version (RELEASING.md): the
// version's shape, package-lock.json, the CHANGELOG section, the README badge and, with --dist,
// the bundled icon. With --notes, writes the GitHub Release notes built from the CHANGELOG.
//
// Usage:
//   node .github/scripts/check-release.mjs                       # before the build
//   node .github/scripts/check-release.mjs --dist                # after the build
//   node .github/scripts/check-release.mjs --dist --notes notes.md --dist-tag beta --repository owner/repo
import { appendFileSync, existsSync, readFileSync, writeFileSync } from "node:fs";
import { checkReleaseFiles, expectedIconSource, releaseNotes } from "./release-checks.mjs";

const args = process.argv.slice(2);
const option = (name) => {
	const i = args.indexOf(name);
	return i >= 0 ? args[i + 1] : undefined;
};
const read = (path) => readFileSync(path, "utf8");

const pkg = JSON.parse(read("package.json"));
const problems = checkReleaseFiles({
	pkg,
	lock: JSON.parse(read("package-lock.json")),
	changelog: read("CHANGELOG.md"),
	readme: read("README.md"),
});

if (args.includes("--dist")) {
	const source = expectedIconSource(pkg.version);
	if (!existsSync("dist/icons/icon.svg")) problems.push("dist/icons/icon.svg is missing: run the build first.");
	else if (read("dist/icons/icon.svg") !== read(source)) {
		problems.push(`dist/icons/icon.svg isn't ${source}, the icon for ${pkg.version}.`);
	}
}

if (problems.length > 0) {
	for (const problem of problems) console.log(`::error::${problem}`);
	process.exit(1);
}
console.log(`✅ release metadata matches ${pkg.version}${args.includes("--dist") ? ", with the right icon" : ""}`);

const notesPath = option("--notes");
if (notesPath) {
	const notes = releaseNotes({
		pkg,
		changelog: read("CHANGELOG.md"),
		distTag: option("--dist-tag") ?? "beta",
		repository: option("--repository") ?? "owner/repo",
	});
	writeFileSync(notesPath, notes);
	// A preview in the run summary, so a rehearsal shows the notes the release will get
	if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, `## Release notes preview\n\n${notes}\n`);
	console.log(`Release notes written to ${notesPath}`);
}
