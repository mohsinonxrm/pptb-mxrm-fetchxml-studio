import { describe, expect, it } from "vitest";
import { cellText, columnToText, copyColumnsFor, rowsToTsv, tsvField, type CopyColumn } from "./copyResults";

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
