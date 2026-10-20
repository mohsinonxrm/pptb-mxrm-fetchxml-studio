/// <reference types="node" />
// Node types for this file only: the tests run the generated quoting through a real shell.
import { execFileSync } from "node:child_process";
import { describe, expect, it } from "vitest";
import {
	generateFxsSyncCode,
	quoteForPosixShell,
	quoteForPowerShell,
} from "./fetchxmlCodeGenerators";

const HOSTILE = `<fetch><entity name="account"><filter><condition attribute="name" operator="like" value="O'Brien O’Neil ‘x’; Remove-Item ‚z‛ $100 $(echo pwned) \`id\` \\ %" /></filter></entity></fetch>`;

/** Whether a POSIX shell is available to run the quoted value through (CI on Linux, Git Bash locally) */
function posixShell(): string | undefined {
	for (const sh of ["sh", "bash"]) {
		try {
			execFileSync(sh, ["-c", "true"], { stdio: "ignore" });
			return sh;
		} catch {
			// try the next one
		}
	}
	return undefined;
}

describe("pac CLI snippet (#166)", () => {
	it("uses pac env fetch with a bash / zsh line and a PowerShell line", () => {
		const code = generateFxsSyncCode("pac", `<fetch>\n  <entity name="account" />\n</fetch>`);
		const lines = code.split("\n").filter((l) => l.startsWith("pac "));
		expect(lines).toEqual([
			`pac env fetch --xml '<fetch><entity name="account" /></fetch>'`,
			`pac env fetch --xml '<fetch><entity name="account" /></fetch>'`,
		]);
		expect(code).toContain("# bash / zsh");
		expect(code).toContain("# PowerShell 7.3 or later");
		expect(code).toContain("--xmlFile");
		expect(code).not.toContain("pac org");
	});

	it("never leaves double quotes around the XML", () => {
		const code = generateFxsSyncCode("pac", HOSTILE);
		expect(code).not.toMatch(/--xml "/);
	});

	it("quotes apostrophes for each shell", () => {
		expect(quoteForPosixShell("O'Brien")).toBe(`'O'\\''Brien'`);
		expect(quoteForPowerShell("O'Brien")).toBe(`'O''Brien'`);
		// PowerShell treats the curly single quotes U+2018–U+201B as single quotes (#196)
		expect(quoteForPowerShell("O’Neil")).toBe("'O’’Neil'");
		expect(quoteForPowerShell("‘x’ ‚y‛")).toBe("'‘‘x’’ ‚‚y‛‛'");
	});

	it.runIf(posixShell())(
		"a POSIX shell receives the XML byte for byte, with nothing expanded",
		() => {
			const sh = posixShell()!;
			const out = execFileSync(sh, ["-c", `printf '%s' ${quoteForPosixShell(HOSTILE)}`], {
				encoding: "utf8",
			});
			expect(out).toBe(HOSTILE);
		},
	);

	it.runIf(hasPwsh())("PowerShell receives the XML unchanged, with nothing expanded", () => {
		const out = execFileSync(
			"pwsh",
			["-NoProfile", "-Command", `[Console]::OutputEncoding = [Text.Encoding]::UTF8; [Console]::Out.Write(${quoteForPowerShell(HOSTILE)})`],
			{
				encoding: "utf8",
			},
		);
		expect(out).toBe(HOSTILE);
	});
});

function hasPwsh(): boolean {
	try {
		execFileSync("pwsh", ["-NoProfile", "-Command", "exit 0"], { stdio: "ignore" });
		return true;
	} catch {
		return false;
	}
}
