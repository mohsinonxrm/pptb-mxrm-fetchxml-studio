import { describe, expect, it } from "vitest";
import { badgeFor, changelogSections, checkReleaseFiles, expectedIconSource, releaseNotes } from "./release-checks.mjs";

const pkg = { name: "@scope/tool", displayName: "Tool Studio", version: "1.5.0-beta.18", features: { minAPI: "1.2.7" } };
const lock = { version: "1.5.0-beta.18", packages: { "": { version: "1.5.0-beta.18" } } };
const changelog = `# Changelog

## [Unreleased]

## [1.5.0-beta.18] - 2026-11-01

### 🔧 Changed

- **A change (#69).** Details.

## [1.5.0-beta.17] - 2026-10-31

- Older.
`;
const readme = `![Version](${badgeFor("1.5.0-beta.18")})`;
const files = { pkg, lock, changelog, readme };

describe("checkReleaseFiles", () => {
	it("passes when everything matches the version", () => {
		expect(checkReleaseFiles(files)).toEqual([]);
	});

	it("allows entries under [Unreleased] above the release", () => {
		const withPending = changelog.replace("## [Unreleased]\n", "## [Unreleased]\n\n### 🔧 Fixed\n\n- **Pending.**\n");
		expect(checkReleaseFiles({ ...files, changelog: withPending })).toEqual([]);
	});

	it("flags a version release.yml doesn't publish", () => {
		const v = "1.5.0-alpha.1";
		const problems = checkReleaseFiles({ ...files, pkg: { ...pkg, version: v } });
		expect(problems[0]).toMatch(/isn't X\.Y\.Z/);
	});

	it("flags package-lock.json left on the old version", () => {
		const problems = checkReleaseFiles({ ...files, lock: { ...lock, packages: { "": { version: "1.5.0-beta.17" } } } });
		expect(problems).toEqual([expect.stringMatching(/^package-lock\.json has 1\.5\.0-beta\.18 \/ 1\.5\.0-beta\.17/)]);
	});

	it("flags a CHANGELOG whose newest section isn't the version", () => {
		const stale = changelog.replace("## [1.5.0-beta.18] - 2026-11-01", "## [Unreleased-old]");
		expect(checkReleaseFiles({ ...files, changelog: stale }).join(" ")).toMatch(/newest release section is \[Unreleased-old\], not \[1\.5\.0-beta\.18\]/);
	});

	it("flags a missing [Unreleased], a missing or invalid date and an empty section", () => {
		expect(checkReleaseFiles({ ...files, changelog: changelog.replace("## [Unreleased]\n", "") }).join(" ")).toMatch(/must start with/);
		expect(checkReleaseFiles({ ...files, changelog: changelog.replace(" - 2026-11-01", "") }).join(" ")).toMatch(/needs a date/);
		expect(checkReleaseFiles({ ...files, changelog: changelog.replace("2026-11-01", "2026-02-30") }).join(" ")).toMatch(/needs a date/);
		const empty = changelog.replace("### 🔧 Changed\n\n- **A change (#69).** Details.\n", "");
		expect(checkReleaseFiles({ ...files, changelog: empty }).join(" ")).toMatch(/has no entries/);
	});

	it("flags a duplicated section", () => {
		const twice = `${changelog}\n## [1.5.0-beta.18] - 2026-11-01\n\n- Again.\n`;
		expect(checkReleaseFiles({ ...files, changelog: twice }).join(" ")).toMatch(/2 \[1\.5\.0-beta\.18\] sections/);
	});

	it("flags a stale README badge", () => {
		expect(checkReleaseFiles({ ...files, readme: `![Version](${badgeFor("1.5.0-beta.17")})` })).toEqual([
			expect.stringMatching(/badge isn't 1\.5\.0-beta\.18/),
		]);
	});
});

describe("badgeFor / expectedIconSource / changelogSections", () => {
	it("escapes dashes for shields.io", () => {
		expect(badgeFor("1.5.0-beta.18")).toBe("https://img.shields.io/badge/version-1.5.0--beta.18-blue");
		expect(badgeFor("1.4.6")).toBe("https://img.shields.io/badge/version-1.4.6-blue");
	});

	it("bundles the Insider icon for pre-releases only", () => {
		expect(expectedIconSource("1.5.0-beta.18")).toBe("icons/icon-insider.svg");
		expect(expectedIconSource("1.5.0-rc.1")).toBe("icons/icon-insider.svg");
		expect(expectedIconSource("1.5.0")).toBe("icons/icon.svg");
	});

	it("reads headings and bodies, normalizing CRLF line endings", () => {
		const sections = changelogSections(changelog.replace(/\n/g, "\r\n"));
		expect(sections.map((s) => [s.name, s.date])).toEqual([
			["Unreleased", undefined],
			["1.5.0-beta.18", "2026-11-01"],
			["1.5.0-beta.17", "2026-10-31"],
		]);
		expect(sections[1].body).toBe("### 🔧 Changed\n\n- **A change (#69).** Details.");
	});
});

describe("releaseNotes", () => {
	it("leads with the channel and the PPTB requirement, then the CHANGELOG section", () => {
		const notes = releaseNotes({ pkg, changelog, distTag: "beta", repository: "o/r" });
		expect(notes).toBe(`**Beta: preview of Tool Studio 1.5.0.** Install it from npm \`beta\`; the stable release stays on npm \`latest\`. The gold Insider icon marks this build in the ToolBox.

> ⚠️ **Requires Power Platform ToolBox 1.2.7 or later** (\`minAPI\` 1.2.7).

### 🔧 Changed

- **A change (#69).** Details.

See [CHANGELOG.md](https://github.com/o/r/blob/v1.5.0-beta.18/CHANGELOG.md) for earlier versions.
`);
	});

	it("words stable and support releases for their channel", () => {
		const stable = { ...pkg, version: "1.5.0" };
		const cl = changelog.replace("[1.5.0-beta.18]", "[1.5.0]");
		expect(releaseNotes({ pkg: stable, changelog: cl, distTag: "latest", repository: "o/r" })).toMatch(/^\*\*Tool Studio 1\.5\.0\.\*\* Install or update it from npm `latest`\./);
		expect(releaseNotes({ pkg: stable, changelog: cl, distTag: "support-1.5", repository: "o/r" })).toMatch(/support release published to npm `support-1\.5`/);
	});

	it("fails without the version's section", () => {
		expect(() => releaseNotes({ pkg: { ...pkg, version: "9.9.9" }, changelog, distTag: "beta", repository: "o/r" })).toThrow(/no \[9\.9\.9\] section/);
	});
});
