import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { exportToExcelLocal } from "./excelExport";
import type { AttributeMetadata } from "./pptbClient";

// Local Excel export: write the workbook, then read it back with ExcelJS (#145).

const FV = "@OData.Community.Display.V1.FormattedValue";

function attr(LogicalName: string, AttributeType: string): AttributeMetadata {
	return { LogicalName, SchemaName: LogicalName, AttributeType, MetadataId: LogicalName };
}

const attributeMetadata = new Map([
	[
		"account",
		new Map(
			[
				attr("name", "String"),
				attr("revenue", "Money"),
				attr("numberofemployees", "Integer"),
				attr("createdon", "DateTime"),
				attr("statuscode", "Status"),
				attr("primarycontactid", "Lookup"),
			].map((a) => [a.LogicalName, a]),
		),
	],
]);

const columns = [
	"name",
	"revenue",
	"numberofemployees",
	"createdon",
	"statuscode",
	"_primarycontactid_value",
];

const columnDisplayNames = new Map([
	["name", "Account Name"],
	["revenue", "Annual Revenue"],
	["numberofemployees", "Employees"],
	["createdon", "Created On"],
	["statuscode", "Status Reason"],
	["_primarycontactid_value", "Primary Contact"],
]);

const records: Record<string, unknown>[] = [
	{
		name: "Contoso",
		revenue: 150000,
		[`revenue${FV}`]: "$150,000.00",
		numberofemployees: 42,
		createdon: "2026-10-01T09:00:00Z",
		[`createdon${FV}`]: "10/1/2026 9:00 AM",
		statuscode: 1,
		[`statuscode${FV}`]: "Active",
		_primarycontactid_value: "00000000-0000-0000-0000-000000000001",
		[`_primarycontactid_value${FV}`]: "Yvonne McKay",
	},
	{ name: "Fabrikam", revenue: null, numberofemployees: null, createdon: null, statuscode: 2 },
];

async function readBack(buffer: ArrayBuffer): Promise<ExcelJS.Worksheet> {
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.load(buffer);
	const sheet = workbook.getWorksheet("Data");
	if (!sheet) throw new Error("Data sheet missing");
	return sheet;
}

describe("exportToExcelLocal", () => {
	it("writes typed, formatted cells that ExcelJS reads back", async () => {
		const { buffer, fileName } = await exportToExcelLocal({
			records,
			columns,
			columnDisplayNames,
			attributeMetadata,
			entityName: "account",
			fileName: "accounts",
		});

		expect(fileName).toMatch(/^accounts_\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}\.xlsx$/);
		const sheet = await readBack(buffer);

		expect(sheet.getRow(1).values).toEqual([
			undefined,
			"Account Name",
			"Annual Revenue",
			"Employees",
			"Created On",
			"Status Reason",
			"Primary Contact",
		]);
		expect(sheet.getRow(1).font?.bold).toBe(true);

		const row = sheet.getRow(2);
		expect(row.getCell(1).value).toBe("Contoso");
		expect(row.getCell(2).value).toBe(150000);
		expect(row.getCell(2).numFmt).toBe('"$"#,##0.00');
		expect(row.getCell(3).value).toBe(42);
		expect(row.getCell(3).numFmt).toBe("#,##0");
		expect(row.getCell(4).value).toEqual(new Date("2026-10-01T09:00:00Z"));
		expect(row.getCell(4).numFmt).toBe("yyyy-mm-dd hh:mm:ss");
		expect(row.getCell(5).value).toBe("Active");
		expect(row.getCell(6).value).toBe("Yvonne McKay");

		// Empty values are written as empty text; a choice without a label keeps its raw value
		const second = sheet.getRow(3);
		expect(second.getCell(2).value).toBe("");
		expect(second.getCell(5).value).toBe(2);

		expect(sheet.rowCount).toBe(3);
		expect(sheet.views[0]).toMatchObject({ state: "frozen", ySplit: 1 });
		expect(sheet.autoFilter).toBe("A1:F3");
	});

	it("adds a (Raw) column next to formatted ones in 'both' mode, with logical-name headers", async () => {
		const { buffer } = await exportToExcelLocal({
			records,
			columns: ["name", "statuscode"],
			columnDisplayNames,
			attributeMetadata,
			entityName: "account",
			displaySettings: {
				useLogicalNames: true,
				valueDisplayMode: "both",
				entityScopeMode: "all",
				advancedFindOnly: true,
				scopeCollapsed: false,
				gridDensity: "medium",
				treeIconStyle: "tinted",
			},
		});
		const sheet = await readBack(buffer);

		expect(sheet.getRow(1).values).toEqual([
			undefined,
			"name",
			"statuscode",
			"statuscode (Raw)",
		]);
		expect(sheet.getRow(2).values).toEqual([undefined, "Contoso", "Active", 1]);
	});
});

// exceljs only calls uuid (v4) for conditional-format rules that need an x14 extension.
// package.json overrides exceljs's uuid (#145); this checks that path still works in Node.
// (FXS's browser build uses exceljs's prebuilt bundle, which carries its own uuid copy.)
describe("exceljs with the uuid override", () => {
	it("writes and reads back a data bar's x14 id", async () => {
		const workbook = new ExcelJS.Workbook();
		const sheet = workbook.addWorksheet("Data");
		sheet.addRows([[1], [2]]);
		sheet.addConditionalFormatting({
			ref: "A1:A2",
			rules: [
				{
					type: "dataBar",
					priority: 1,
					cfvo: [{ type: "min" }, { type: "max" }],
					color: { argb: "FF638EC6" },
					gradient: false,
				} as unknown as ExcelJS.ConditionalFormattingRule,
			],
		});
		const buffer = await workbook.xlsx.writeBuffer();

		const reloaded = await readBack(buffer as ArrayBuffer);
		// Not in exceljs's typings, but part of the worksheet model it reads back
		const { conditionalFormattings } = reloaded as unknown as {
			conditionalFormattings: { rules: { type: string; x14Id?: string }[] }[];
		};
		const rule = conditionalFormattings[0].rules[0];
		expect(rule.type).toBe("dataBar");
		expect(rule.x14Id).toMatch(
			/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/,
		);
	});
});
