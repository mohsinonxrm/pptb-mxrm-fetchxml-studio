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
import { generateFxsSyncCode, generatePowerAutomateSpec, type FxsSyncFormat } from "./fetchxmlCodeGenerators";

const ORG_URL = "https://contoso.crm.dynamics.com";
const FORMATS: FxsSyncFormat[] = ["csharp-fetchexpr", "javascript", "pac", "webapi"];
const EXAMPLES = LEARN_FETCHXML_EXAMPLES.map((example, i) => [i + 1, example.xml] as const);

describe("golden examples", () => {
	it("cover every Learn example", () => {
		expect(EXAMPLES.length).toBeGreaterThanOrEqual(60);
	});
});

describe.each(FORMATS)("%s", (format) => {
	it.each(EXAMPLES)("Learn example #%i", (_index, xml) => {
		expect(generateFxsSyncCode(format, xml, undefined, ORG_URL)).toMatchSnapshot();
	});
});

describe("Power Automate List rows spec", () => {
	it.each(EXAMPLES)("Learn example #%i", (_index, xml) => {
		expect(generatePowerAutomateSpec(xml)).toMatchSnapshot();
	});
});
