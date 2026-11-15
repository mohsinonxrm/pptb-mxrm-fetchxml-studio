// README badges ToolBox would show as raw Markdown (#348): the release check catches them
import { describe, expect, it } from "vitest";
import { badgeMarkdown, badges } from "./badges.mjs";
import { checkReleaseFiles } from "./release-checks.mjs";

const pkg = {
	name: "@scope/tool",
	version: "1.5.1",
	license: "AGPL-3.0-only",
	repository: { type: "git", url: "git+https://github.com/owner/tool.git" },
	features: { minAPI: "1.2.7" },
	dependencies: { react: "^18.3.1", "@fluentui/react-components": "^9.74.9" },
	devDependencies: { typescript: "~5.9.3", "@pptb/types": "1.2.7" },
};
const lock = { version: "1.5.1", packages: { "": { version: "1.5.1" } } };
const changelog = "# Changelog\n\n## [Unreleased]\n\n## [1.5.1] - 2026-11-15\n\n- **A fix.**\n";
const badgeFiles = Object.fromEntries(badges(pkg).map((b) => [b.file, b.svg]));

describe("README badges that ToolBox shows as raw Markdown (#348)", () => {
	it("flags the version badge wrapped in a link (the 1.5.0 form)", () => {
		const [version, ...rest] = badgeMarkdown(pkg);
		const readme = `# Tool\n\n[${version}](https://www.npmjs.com/package/@scope/tool)\n${rest.join("\n")}\n`;
		expect(checkReleaseFiles({ pkg, lock, changelog, readme, badgeFiles })).not.toEqual([]);
	});
});
