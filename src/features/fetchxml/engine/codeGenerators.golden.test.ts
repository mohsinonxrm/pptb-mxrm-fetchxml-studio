/**
 * Golden output for the local code generators (#189): every FetchXML example from the Microsoft
 * Learn articles (`learnFetchXml.ts`) through C# FetchExpression, JavaScript, pac CLI, Web API and
 * the Power Automate "List rows" spec. Any change to a snippet shows up as a snapshot diff, so it's
 * reviewed rather than slipping through a refactor. C# QueryExpression comes from the Web API's
 * FetchXmlToQueryExpression response and has its own tests (queryExpressionCodegen.test.ts).
 *
 * After an intended change: `npx vitest run -u src/features/fetchxml/engine/codeGenerators.golden.test.ts`,
 * then review the snapshot diff.
 */
import { describe, expect, it } from "vitest";
import { LEARN_FETCHXML_EXAMPLES } from "../model/__fixtures__/learnFetchXml";
import {
	generateFxsSyncCode,
	generatePowerAutomateSpec,
	type FxsSyncFormat,
} from "./fetchxmlCodeGenerators";

const ORG_URL = "https://contoso.crm.dynamics.com";
const FORMATS: FxsSyncFormat[] = ["csharp-fetchexpr", "javascript", "pac", "webapi"];

// Named by the Learn article and the example's place in it, so adding an example only renames the
// cases of its own article
const perArticle = new Map<string, number>();
const EXAMPLES = LEARN_FETCHXML_EXAMPLES.map((example) => {
	const article = example.source.split("/").filter(Boolean).pop() ?? "example";
	const n = (perArticle.get(article) ?? 0) + 1;
	perArticle.set(article, n);
	return [`${article} #${n}`, example.xml] as const;
});

describe("golden examples", () => {
	it("cover every Learn example", () => {
		expect(EXAMPLES.length).toBeGreaterThanOrEqual(60);
		expect(new Set(EXAMPLES.map(([name]) => name)).size).toBe(EXAMPLES.length);
	});
});

describe.each(FORMATS)("%s", (format) => {
	it.each(EXAMPLES)("%s", (_name, xml) => {
		expect(generateFxsSyncCode(format, xml, undefined, ORG_URL)).toMatchSnapshot();
	});
});

describe("Power Automate List rows spec", () => {
	it.each(EXAMPLES)("%s", (_name, xml) => {
		expect(generatePowerAutomateSpec(xml)).toMatchSnapshot();
	});
});

// The Code panel passes the entity set name from metadata when it has it (CodePanel.tsx), which
// wins over the plural derived from the entity name
describe("with the entity set name from metadata", () => {
	const xml =
		"<fetch top='5'><entity name='adminsettingsentity'><attribute name='name'/></entity></fetch>";
	it("Web API uses it in the URL", () => {
		const code = generateFxsSyncCode("webapi", xml, "new_admin_settings_entities", ORG_URL);
		expect(code).toContain("/api/data/v9.2/new_admin_settings_entities?fetchXml=");
		expect(code).toMatchSnapshot();
	});
	it("Power Automate uses it as the table name", () => {
		const spec = generatePowerAutomateSpec(xml, "new_admin_settings_entities");
		expect(spec.fields.find((f) => f.label === "Table name")?.value).toBe(
			"new_admin_settings_entities",
		);
		expect(spec).toMatchSnapshot();
	});
});
