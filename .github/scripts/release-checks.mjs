// What every release must update, checked against package.json's version (RELEASING.md):
// the version's shape, package-lock.json, the CHANGELOG section, the README badge and the bundled
// icon. Also builds the GitHub Release notes from the CHANGELOG section. Used by check-release.mjs;
// pure functions, so they're unit tested (release-checks.test.mjs).

/** X.Y.Z, X.Y.Z-beta.N or X.Y.Z-rc.N: the shapes release.yml publishes */
export const VERSION_PATTERN = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(-(beta|rc)\.(0|[1-9]\d*))?$/;

export const isPrerelease = (version) => version.includes("-");

/** The shields.io badge path for a version: `-` is written `--` */
export const badgeFor = (version) => `https://img.shields.io/badge/version-${version.replace(/-/g, "--")}-blue`;

/** The CHANGELOG's `## [name] - date` headings, in order, with their bodies */
export function changelogSections(changelog) {
	const lines = changelog.split(/\r?\n/);
	const sections = [];
	for (const line of lines) {
		const heading = /^## \[([^\]]+)\](?: - (\S+))?\s*$/.exec(line);
		if (heading) sections.push({ name: heading[1], date: heading[2], body: [] });
		else if (sections.length > 0) sections[sections.length - 1].body.push(line);
	}
	return sections.map((s) => ({ ...s, body: s.body.join("\n").trim() }));
}

const isIsoDate = (date) => /^\d{4}-\d{2}-\d{2}$/.test(date ?? "") && !Number.isNaN(Date.parse(`${date}T00:00:00Z`)) && new Date(`${date}T00:00:00Z`).toISOString().startsWith(date);

/**
 * Problems with the release metadata, or [] when it's consistent.
 * @param {{ pkg: any, lock: any, changelog: string, readme: string }} files
 */
export function checkReleaseFiles({ pkg, lock, changelog, readme }) {
	const problems = [];
	const version = pkg.version;
	if (!VERSION_PATTERN.test(version ?? "")) {
		problems.push(`package.json version "${version}" isn't X.Y.Z, X.Y.Z-beta.N or X.Y.Z-rc.N.`);
	}
	if (lock?.version !== version || lock?.packages?.[""]?.version !== version) {
		problems.push(
			`package-lock.json has ${lock?.version} / ${lock?.packages?.[""]?.version}, not ${version}. Set the version with \`npm version ${version} --no-git-tag-version\`.`,
		);
	}
	const sections = changelogSections(changelog);
	if (sections[0]?.name !== "Unreleased") {
		problems.push("CHANGELOG.md must start with a `## [Unreleased]` section.");
	}
	const released = sections.find((s) => s.name !== "Unreleased");
	if (released?.name !== version) {
		problems.push(
			`CHANGELOG.md's newest release section is [${released?.name}], not [${version}]. Rename [Unreleased] to \`## [${version}] - YYYY-MM-DD\` and add an empty [Unreleased] above it.`,
		);
	} else {
		if (!isIsoDate(released.date)) problems.push(`CHANGELOG.md's [${version}] section needs a date: \`## [${version}] - YYYY-MM-DD\`.`);
		if (!/^\s*- /m.test(released.body)) problems.push(`CHANGELOG.md's [${version}] section has no entries.`);
	}
	const duplicates = sections.filter((s) => s.name === version).length;
	if (duplicates > 1) problems.push(`CHANGELOG.md has ${duplicates} [${version}] sections.`);
	if (!readme.includes(badgeFor(version))) {
		problems.push(`README.md's version badge isn't ${version}: expected ${badgeFor(version)}.`);
	}
	return problems;
}

/** The icon the build must bundle at dist/icons/icon.svg (vite.config.ts toolIcon) */
export const expectedIconSource = (version) => (isPrerelease(version) ? "icons/icon-insider.svg" : "icons/icon.svg");

/**
 * The GitHub Release notes: a line on the channel and the PPTB requirement, then the version's
 * CHANGELOG section. release.yml appends GitHub's generated list of PRs and the compare link.
 * @param {{ pkg: any, changelog: string, distTag: string, repository: string }} input
 */
export function releaseNotes({ pkg, changelog, distTag, repository }) {
	const version = pkg.version;
	const name = pkg.displayName ?? pkg.name;
	const section = changelogSections(changelog).find((s) => s.name === version);
	if (!section) throw new Error(`CHANGELOG.md has no [${version}] section.`);
	const minApi = pkg.features?.minAPI;
	let lead;
	if (distTag === "beta") {
		lead = `**Beta: preview of ${name} ${version.replace(/-.*/, "")}.** Install it from npm \`beta\`; the stable release stays on npm \`latest\`. The gold Insider icon marks this build in the ToolBox.`;
	} else if (distTag === "latest") {
		lead = `**${name} ${version}.** Install or update it from npm \`latest\`.`;
	} else {
		lead = `**${name} ${version}**, a support release published to npm \`${distTag}\`. The current release is on npm \`latest\`.`;
	}
	const parts = [lead];
	if (minApi) parts.push(`> ⚠️ **Requires Power Platform ToolBox ${minApi} or later** (\`minAPI\` ${minApi}).`);
	parts.push(section.body);
	parts.push(`See [CHANGELOG.md](https://github.com/${repository}/blob/v${version}/CHANGELOG.md) for earlier versions.`);
	return `${parts.join("\n\n")}\n`;
}
