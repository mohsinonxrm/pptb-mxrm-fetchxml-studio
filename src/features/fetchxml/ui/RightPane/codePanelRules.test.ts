import { describe, expect, it } from "vitest";
import {
	CODE_LANGUAGES,
	CODE_LANGUAGE_LABELS,
	copyAllText,
	copyDisabledReason,
	needsConnection,
	shownLanguage,
	sourceNote,
} from "./codePanelRules";

describe("the Code tab's languages (#233, decision 3)", () => {
	it("are in the owner's order, with these labels", () => {
		expect(CODE_LANGUAGES.map((l) => CODE_LANGUAGE_LABELS[l])).toEqual([
			"C# QueryExpression",
			"C# FetchExpression",
			"Web API",
			"Xrm.WebApi",
			"Power Automate",
			"pac CLI",
			"SQL (TDS)",
		]);
	});

	it("only C# QueryExpression and SQL need a connection, and name their Dataverse function", () => {
		expect(CODE_LANGUAGES.filter(needsConnection)).toEqual(["csharp-queryexpr", "sql"]);
		expect(sourceNote("csharp-queryexpr")).toBe("From FetchXmlToQueryExpression");
		expect(sourceNote("sql")).toBe("From FetchXMLToSQL");
		expect(sourceNote("csharp-fetchexpr")).toBeUndefined();
	});
});

describe("the language shown (decisions 2 and 4)", () => {
	it("is the one picked this session", () => {
		expect(shownLanguage("pac", true)).toBe("pac");
		expect(shownLanguage("csharp-queryexpr", false)).toBe("csharp-queryexpr");
	});
	it("is C# QueryExpression by default, C# FetchExpression while not connected", () => {
		expect(shownLanguage(null, true)).toBe("csharp-queryexpr");
		expect(shownLanguage(null, false)).toBe("csharp-fetchexpr");
	});
});

describe("why Copy is off", () => {
	const ok = { connected: true, loading: false, failed: false, text: "code" };
	it("in the order the user would fix it", () => {
		expect(copyDisabledReason("sql", { ...ok, connected: false, loading: true })).toBe(
			"Connect to an environment to generate SQL (TDS)",
		);
		expect(copyDisabledReason("csharp-queryexpr", { ...ok, loading: true })).toBe(
			"Generating…",
		);
		expect(copyDisabledReason("csharp-queryexpr", { ...ok, failed: true })).toBe(
			"Nothing to copy: generation failed",
		);
		expect(copyDisabledReason("pac", { ...ok, text: "" })).toBe("There's no code to copy yet");
		expect(copyDisabledReason("powerautomate", { ...ok, text: "" })).toBe(
			"Nothing to copy: no field has a value yet",
		);
	});
	it("not for a language made in the tool, offline", () => {
		expect(copyDisabledReason("csharp-fetchexpr", { ...ok, connected: false })).toBeUndefined();
	});
});

describe("Power Automate's Copy all", () => {
	it("copies the filled-in fields, label then value, one per line", () => {
		expect(
			copyAllText({
				actionName: "List rows",
				connector: "Microsoft Dataverse connector",
				fields: [
					{ label: "Table name", value: "accounts" },
					{ label: "Filter rows", value: "" },
					{ label: "Fetch Xml Query", value: "<fetch/>" },
				],
			}),
		).toBe(`${"Table name".padEnd(28)} accounts\n${"Fetch Xml Query".padEnd(28)} <fetch/>`);
	});
});
