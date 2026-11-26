import { describe, expect, it } from "vitest";
import {
	badgeMarkdown,
	badges,
	badgeSvg,
	badgeUrl,
	readmeRef,
	readmeUrlFor,
	rewriteBadgeLines,
	rewriteReadmeUrl,
	versionRef,
} from "./badges.mjs";
import {
	changelogSections,
	checkReleaseFiles,
	expectedIconSource,
	releaseNotes,
} from "./release-checks.mjs";

const pkg = {
	name: "@scope/tool",
	displayName: "Tool Studio",
	version: "1.5.0-beta.18",
	license: "AGPL-3.0-only",
	repository: { type: "git", url: "git+https://github.com/owner/tool.git" },
	features: { minAPI: "1.2.7" },
	dependencies: { react: "^18.3.1", "@fluentui/react-components": "^9.74.9" },
	devDependencies: { typescript: "~5.9.3", "@pptb/types": "1.2.7" },
	configurations: {
		readmeUrl: "https://raw.githubusercontent.com/owner/tool/refs/heads/release/1.5/README.md",
	},
};
const lock = { version: "1.5.0-beta.18", packages: { "": { version: "1.5.0-beta.18" } } };
const changelog = `# Changelog

## [Unreleased]

## [1.5.0-beta.18] - 2026-11-01

### 🔧 Changed

- **A change (#69).** Details.

## [1.5.0-beta.17] - 2026-10-31

- Older.
`;
const readme = `# Tool

${badgeMarkdown(pkg).join("\n")}
`;
const badgeFiles = Object.fromEntries(badges(pkg).map((b) => [b.file, b.svg]));
const files = { pkg, lock, changelog, readme, badgeFiles };

describe("checkReleaseFiles", () => {
	it("passes when everything matches the version", () => {
		expect(checkReleaseFiles(files)).toEqual([]);
	});

	it("allows entries under [Unreleased] above the release", () => {
		const withPending = changelog.replace(
			"## [Unreleased]\n",
			"## [Unreleased]\n\n### 🔧 Fixed\n\n- **Pending.**\n",
		);
		expect(checkReleaseFiles({ ...files, changelog: withPending })).toEqual([]);
	});

	it("flags a version release.yml doesn't publish", () => {
		const v = "1.5.0-alpha.1";
		const problems = checkReleaseFiles({ ...files, pkg: { ...pkg, version: v } });
		expect(problems[0]).toMatch(/isn't X\.Y\.Z/);
	});

	it("flags package-lock.json left on the old version", () => {
		const problems = checkReleaseFiles({
			...files,
			lock: { ...lock, packages: { "": { version: "1.5.0-beta.17" } } },
		});
		expect(problems).toEqual([
			expect.stringMatching(/^package-lock\.json has 1\.5\.0-beta\.18 \/ 1\.5\.0-beta\.17/),
		]);
	});

	it("flags a CHANGELOG whose newest section isn't the version", () => {
		const stale = changelog.replace("## [1.5.0-beta.18] - 2026-11-01", "## [Unreleased-old]");
		expect(checkReleaseFiles({ ...files, changelog: stale }).join(" ")).toMatch(
			/newest release section is \[Unreleased-old\], not \[1\.5\.0-beta\.18\]/,
		);
	});

	it("flags a missing [Unreleased], a missing or invalid date and an empty section", () => {
		expect(
			checkReleaseFiles({
				...files,
				changelog: changelog.replace("## [Unreleased]\n", ""),
			}).join(" "),
		).toMatch(/must start with/);
		expect(
			checkReleaseFiles({ ...files, changelog: changelog.replace(" - 2026-11-01", "") }).join(
				" ",
			),
		).toMatch(/needs a date/);
		expect(
			checkReleaseFiles({
				...files,
				changelog: changelog.replace("2026-11-01", "2026-02-30"),
			}).join(" "),
		).toMatch(/needs a date/);
		const empty = changelog.replace("### 🔧 Changed\n\n- **A change (#69).** Details.\n", "");
		expect(checkReleaseFiles({ ...files, changelog: empty }).join(" ")).toMatch(
			/has no entries/,
		);
	});

	it("flags a duplicated section", () => {
		const twice = `${changelog}\n## [1.5.0-beta.18] - 2026-11-01\n\n- Again.\n`;
		expect(checkReleaseFiles({ ...files, changelog: twice }).join(" ")).toMatch(
			/2 \[1\.5\.0-beta\.18\] sections/,
		);
	});

	it("flags badge files not regenerated for the version, or missing", () => {
		const old = badges({ ...pkg, version: "1.5.0-beta.17" }).find(
			(b) => b.file === "version.svg",
		).svg;
		expect(
			checkReleaseFiles({ ...files, badgeFiles: { ...badgeFiles, "version.svg": old } }),
		).toEqual(["assets/badges/ is out of date (version.svg): run `npm run badges`."]);
		expect(
			checkReleaseFiles({
				...files,
				badgeFiles: { ...badgeFiles, "react.svg": undefined },
			}).join(" "),
		).toMatch(/out of date \(react\.svg\)/);
	});

	it("flags a README without the badges, or with img.shields.io", () => {
		expect(checkReleaseFiles({ ...files, readme: "# Tool\n" }).join(" ")).toMatch(
			/README\.md doesn't show these badges/,
		);
		const shields = `${readme}![npm](https://img.shields.io/npm/v/@scope/tool)\n`;
		expect(checkReleaseFiles({ ...files, readme: shields })).toEqual([
			expect.stringMatching(/img\.shields\.io, which PPTB's README viewer blocks/),
		]);
		expect(
			checkReleaseFiles({
				...files,
				readme: `${readme}The badges aren't on img.shields.io.
`,
			}),
		).toEqual([]);
	});
});

describe("badges / expectedIconSource / changelogSections", () => {
	it("generates every badge from package.json, linked from the version's tag on raw.githubusercontent.com (#379)", () => {
		expect(badges(pkg).map((b) => [b.file, b.alt])).toEqual([
			["version.svg", "Version"],
			["pptb.svg", "Power Platform ToolBox 1.2.7 or later"],
			["pptb-types.svg", "@pptb/types 1.2.7"],
			["react.svg", "React 18"],
			["typescript.svg", "TypeScript 5.9"],
			["fluent-ui.svg", "Fluent UI v9"],
			["license.svg", "License AGPL-3.0-only"],
		]);
		expect(badgeUrl(pkg, "react.svg")).toBe(
			"https://raw.githubusercontent.com/owner/tool/refs/tags/v1.5.0-beta.18/assets/badges/react.svg",
		);
		// The version badge has no version in its alt text; and it isn't a link, which ToolBox's README
		// renderer would show as raw Markdown (#348)
		expect(badgeMarkdown(pkg)[0]).toBe(
			"![Version](https://raw.githubusercontent.com/owner/tool/refs/tags/v1.5.0-beta.18/assets/badges/version.svg)",
		);
	});

	it("colours pre-releases yellow and releases blue, and escapes text", () => {
		const svgOf = (version) =>
			badges({ ...pkg, version }).find((b) => b.file === "version.svg").svg;
		expect(svgOf("1.5.0-beta.18")).toContain('fill="#dfb317"');
		expect(svgOf("1.5.0")).toContain('fill="#007ec6"');
		expect(svgOf("1.5.0")).toContain('aria-label="version: 1.5.0"');
		expect(badgeSvg("a<b", "c&d", "#000")).toContain("a&lt;b: c&amp;d");
	});

	it("bundles the Insider icon for pre-releases only", () => {
		expect(expectedIconSource("1.5.0-beta.18")).toBe("icons/icon-insider.svg");
		expect(expectedIconSource("1.5.0-rc.1")).toBe("icons/icon-insider.svg");
		expect(expectedIconSource("1.5.0")).toBe("icons/icon.svg");
	});

	it("reads the date of headings with text after it, and keeps their bodies apart", () => {
		const extra = changelog.replace(
			"## [1.5.0-beta.17] - 2026-10-31",
			"## [1.5.0-beta.17] - 2026-10-31 [YANKED]",
		);
		const sections = changelogSections(extra);
		expect(sections[2]).toMatchObject({
			name: "1.5.0-beta.17",
			date: "2026-10-31",
			body: "- Older.",
		});
		expect(sections[1].body).not.toContain("Older");
		const stable = changelog.replace(
			"## [1.5.0-beta.18] - 2026-11-01",
			"## [1.5.0-beta.18] - 2026-11-01 (stable)",
		);
		expect(checkReleaseFiles({ ...files, changelog: stable })).toEqual([]);
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
		expect(notes)
			.toBe(`**Beta: preview of Tool Studio 1.5.0.** Install it from npm \`beta\`; the stable release stays on npm \`latest\`. The gold Insider icon marks this build in the ToolBox.

> ⚠️ **Requires Power Platform ToolBox 1.2.7 or later** (\`minAPI\` 1.2.7).

### 🔧 Changed

- **A change (#69).** Details.

See [CHANGELOG.md](https://github.com/o/r/blob/v1.5.0-beta.18/CHANGELOG.md) for earlier versions.
`);
	});

	it("calls an rc a release candidate", () => {
		const rc = { ...pkg, version: "1.5.0-rc.1" };
		const cl = changelog.replace("[1.5.0-beta.18]", "[1.5.0-rc.1]");
		expect(
			releaseNotes({ pkg: rc, changelog: cl, distTag: "beta", repository: "o/r" }),
		).toMatch(/^\*\*Release candidate: preview of Tool Studio 1\.5\.0\.\*\*/);
	});

	it("words stable and support releases for their channel", () => {
		const stable = { ...pkg, version: "1.5.0" };
		const cl = changelog.replace("[1.5.0-beta.18]", "[1.5.0]");
		expect(
			releaseNotes({ pkg: stable, changelog: cl, distTag: "latest", repository: "o/r" }),
		).toMatch(/^\*\*Tool Studio 1\.5\.0\.\*\* Install or update it from npm `latest`\./);
		expect(
			releaseNotes({ pkg: stable, changelog: cl, distTag: "support-1.5", repository: "o/r" }),
		).toMatch(/support release published to npm `support-1\.5`/);
	});

	it("fails without the version's section", () => {
		expect(() =>
			releaseNotes({
				pkg: { ...pkg, version: "9.9.9" },
				changelog,
				distTag: "beta",
				repository: "o/r",
			}),
		).toThrow(/no \[9\.9\.9\] section/);
	});
});

describe("each version's README: badges at its tag, readmeUrl that exists before the tag (#379)", () => {
	const REPO = "https://raw.githubusercontent.com/owner/tool";
	const at = (version) => ({ ...pkg, version });

	it("reads a prerelease's README from its release branch, and a stable release's from main", () => {
		expect(readmeRef(at("1.7.0-beta.1"))).toBe("refs/heads/release/1.7");
		expect(readmeRef(at("1.7.0-rc.2"))).toBe("refs/heads/release/1.7");
		expect(readmeRef(at("1.10.3-beta.4"))).toBe("refs/heads/release/1.10");
		expect(readmeRef(at("1.7.0"))).toBe("refs/heads/main");
		expect(readmeUrlFor(at("1.7.0-beta.1"))).toBe(`${REPO}/refs/heads/release/1.7/README.md`);
		expect(readmeUrlFor(at("1.7.0"))).toBe(`${REPO}/refs/heads/main/README.md`);
	});

	it("links the badges at the version's own tag", () => {
		expect(versionRef(at("1.7.0-beta.1"))).toBe("refs/tags/v1.7.0-beta.1");
		expect(badgeUrl(at("1.7.0"), "pptb.svg")).toBe(
			`${REPO}/refs/tags/v1.7.0/assets/badges/pptb.svg`,
		);
	});

	it("fails the release check when readmeUrl isn't the version's", () => {
		const onMain = {
			...pkg,
			configurations: { readmeUrl: `${REPO}/refs/heads/main/README.md` },
		};
		expect(checkReleaseFiles({ ...files, pkg: onMain }).join(" ")).toMatch(
			/configurations\.readmeUrl isn't https:\/\/raw\.githubusercontent\.com\/owner\/tool\/refs\/heads\/release\/1\.5\/README\.md: run `npm run badges`/,
		);
		const missing = { ...pkg, configurations: undefined };
		expect(checkReleaseFiles({ ...files, pkg: missing }).join(" ")).toMatch(/readmeUrl/);
	});

	it("fails the release check when the README's badges are still on main", () => {
		const onMain = readme.replaceAll("refs/tags/v1.5.0-beta.18", "refs/heads/main");
		expect(checkReleaseFiles({ ...files, readme: onMain }).join(" ")).toMatch(
			/README\.md doesn't show these badges/,
		);
	});

	it("rewrites each badge line, from main or an older tag, and nothing else", () => {
		const next = at("1.7.0-beta.1");
		const files = badges(next).map((b) => b.file);
		const old = badgeMarkdown(at("1.6.0"));
		const onMain = old.map((l) => l.replace("refs/tags/v1.6.0", "refs/heads/main"));
		const screenshot = `![The query tree](${REPO}/refs/heads/main/assets/screenshots/tree.png)`;
		const mention = `Badges live in \`${REPO}/refs/heads/main/assets/badges/version.svg\` and more.`;
		const readmeText = `# Tool\n\n${onMain.slice(0, 3).join("\n")}\n${old.slice(3).join("\n")}\n\n${screenshot}\n\n${mention}\n`;
		const rewritten = rewriteBadgeLines(readmeText, {
			repository: "owner/tool",
			files,
			lines: badgeMarkdown(next),
		});
		expect(rewritten).toBe(
			`# Tool\n\n${badgeMarkdown(next).join("\n")}\n\n${screenshot}\n\n${mention}\n`,
		);
		// Another repository's badge isn't touched
		const other =
			"![Version](https://raw.githubusercontent.com/someone/else/refs/heads/main/assets/badges/version.svg)";
		expect(
			rewriteBadgeLines(other, {
				repository: "owner/tool",
				files,
				lines: badgeMarkdown(next),
			}),
		).toBe(other);
	});

	it("changes only readmeUrl's value in package.json's text, keeping its formatting", () => {
		const text = `{\n\t"name": "x",\n\t"configurations": {\n\t\t"readmeUrl": "${REPO}/refs/heads/main/README.md",\n\t\t"repository": "y"\n\t}\n}\n`;
		expect(rewriteReadmeUrl(text, `${REPO}/refs/heads/release/1.7/README.md`)).toBe(
			text.replace("refs/heads/main", "refs/heads/release/1.7"),
		);
	});
});
