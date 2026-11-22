/**
 * The Code tab's rules (#233), apart from the component so they stay fast-refreshable and testable:
 * the languages in the owner's order with their labels, which one opens first, what Copy does and
 * why it's off, and which Dataverse function a language comes from.
 */
import {
	FXS_SYNC_FORMAT_LABELS,
	type FxsSyncFormat,
	type PowerAutomateActionSpec,
} from "../../engine/fetchxmlCodeGenerators";

export type CodeLanguage = FxsSyncFormat | "csharp-queryexpr" | "sql";

/** The language tabs, in the owner's order (decision 3 on #233) */
export const CODE_LANGUAGES: readonly CodeLanguage[] = [
	"csharp-queryexpr",
	"csharp-fetchexpr",
	"webapi",
	"javascript",
	"powerautomate",
	"pac",
	"sql",
];

export const CODE_LANGUAGE_LABELS: Record<CodeLanguage, string> = {
	...FXS_SYNC_FORMAT_LABELS,
	"csharp-queryexpr": "C# QueryExpression",
	// It's the client API, not the Web API tab's URL / cURL / PowerShell
	javascript: "Xrm.WebApi",
	sql: "SQL (TDS)",
};

/** The languages Dataverse generates: they need a connection */
export const needsConnection = (language: CodeLanguage): boolean =>
	language === "csharp-queryexpr" || language === "sql";

/**
 * The language the tab shows: the one picked this session, else C# QueryExpression, or C#
 * FetchExpression while not connected (decision 4 on #233)
 */
export function shownLanguage(picked: CodeLanguage | null, connected: boolean): CodeLanguage {
	if (picked) return picked;
	return connected ? "csharp-queryexpr" : "csharp-fetchexpr";
}

/** Where a language's code comes from, when it's a Dataverse function */
export function sourceNote(language: CodeLanguage): string | undefined {
	if (language === "csharp-queryexpr") return "From FetchXmlToQueryExpression";
	if (language === "sql") return "From FetchXMLToSQL";
	return undefined;
}

/** Copy's state, in the order the user would fix them */
export interface CopyState {
	connected: boolean;
	loading: boolean;
	failed: boolean;
	/** The code (Power Automate: the filled-in fields as one block) */
	text: string;
}

/** Why Copy (Copy all on Power Automate) can't be used, or undefined */
export function copyDisabledReason(language: CodeLanguage, state: CopyState): string | undefined {
	const label = CODE_LANGUAGE_LABELS[language];
	if (needsConnection(language) && !state.connected)
		return `Connect to an environment to generate ${label}`;
	if (state.loading) return "Generating…";
	if (state.failed) return "Nothing to copy: generation failed";
	if (!state.text)
		return language === "powerautomate"
			? "Nothing to copy: no field has a value yet"
			: "There's no code to copy yet";
	return undefined;
}

/** Power Automate's Copy all: every filled-in field, one per line, label then value */
export function copyAllText(spec: PowerAutomateActionSpec): string {
	return spec.fields
		.filter((f) => f.value)
		.map((f) => `${f.label.padEnd(28)} ${f.value}`)
		.join("\n");
}
