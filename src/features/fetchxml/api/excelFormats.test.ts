import { describe, expect, it } from "vitest";
import { decimalNumberFormat, decimalsShown, groupedNumberFormat, moneyNumberFormat } from "./excelFormats";

describe("Excel number formats (#184)", () => {
	it("decimal precision: 0 is \"0\", never \"0.\"", () => {
		expect(decimalNumberFormat(0)).toBe("0");
		expect(decimalNumberFormat(3)).toBe("0.000");
		expect(decimalNumberFormat(undefined)).toBe("0.00");
		expect(groupedNumberFormat(0)).toBe("#,##0");
		expect(groupedNumberFormat(2)).toBe("#,##0.00");
	});

	it("reads the decimals a formatted amount shows", () => {
		expect(decimalsShown("$150,000.00", 150000)).toBe(2);
		expect(decimalsShown("€1.234,50", 1234.5)).toBe(2);
		expect(decimalsShown("¥1,235", 1235)).toBe(0);
		expect(decimalsShown("1.235 ¥", 1235)).toBe(0);
		expect(decimalsShown("BD 1.234", 1.234)).toBe(3);
		expect(decimalsShown("1,234.567 BD", 1234.567)).toBe(3);
		expect(decimalsShown("$5", 5)).toBe(0);
	});

	it("keeps the currency symbol where Dataverse put it", () => {
		expect(moneyNumberFormat("$150,000.00", 150000)).toBe('"$"#,##0.00');
		expect(moneyNumberFormat("€1.234,50", 1234.5)).toBe('"€"#,##0.00');
		expect(moneyNumberFormat("€ 1.234,50", 1234.5)).toBe('"€ "#,##0.00');
		expect(moneyNumberFormat("1 234,50 kr", 1234.5)).toBe('#,##0.00" kr"');
		expect(moneyNumberFormat("CHF 1'234.50", 1234.5)).toBe("\"CHF \"#,##0.00");
		expect(moneyNumberFormat("¥1,235", 1235)).toBe('"¥"#,##0');
	});

	it("negative amounts: the sign and brackets aren't part of the symbol", () => {
		expect(moneyNumberFormat("-$5.00", -5)).toBe('"$"#,##0.00');
		expect(moneyNumberFormat("($5.00)", -5)).toBe('"$"#,##0.00');
	});

	it("uses the column's precision when it's set, and a plain format without a formatted value", () => {
		expect(moneyNumberFormat("$5.0000", 5, 4)).toBe('"$"#,##0.0000');
		expect(moneyNumberFormat(undefined, 5)).toBe("#,##0.00");
		expect(moneyNumberFormat("", 5)).toBe("#,##0.00");
	});
});
