/**
 * Copying query results to the clipboard (#50)
 *
 * Rows copy as tab-separated values with a header row, so they paste into Excel as cells. Values
 * containing a tab, line break or quote are quoted the way Excel writes them ("a ""b"" c"). Text a
 * spreadsheet would run as a formula gets a leading apostrophe (#208).
 */

import { getFormattedValue } from "./FormattedValueUtils";
import { copyText } from "../../api/host";

/** A grid column as copied: its header and where its value comes from */
export interface CopyColumn {
	/** Grid column ID (a "(Raw)" twin column ends in "__raw") */
	id: string;
	/** Header as the grid shows it */
	header: string;
	/** Result key the value is read from */
	sourceCol: string;
	/** Whether the grid shows this column's raw value */
	showsRaw: boolean;
}

/** "shown": each value as the grid shows it; "raw": raw values (GUIDs, numbers, ISO dates) */
export type CopyValueMode = "shown" | "raw";

const toText = (value: unknown): string => {
	if (value === null || value === undefined) return "";
	if (typeof value === "object") return JSON.stringify(value);
	return String(value);
};

/** A record's raw value for a column, as text ("" when empty) */
export function rawText(record: Record<string, unknown>, sourceCol: string): string {
	return toText(record[sourceCol]);
}

/** A record's formatted value for a column (the raw value when there's no formatted value) */
export function formattedText(record: Record<string, unknown>, sourceCol: string): string {
	return toText(getFormattedValue(record, sourceCol) ?? record[sourceCol]);
}

/** A cell's text: as the grid shows it, or its raw value */
export function cellText(record: Record<string, unknown>, column: CopyColumn, mode: CopyValueMode = "shown"): string {
	return mode === "raw" || column.showsRaw
		? rawText(record, column.sourceCol)
		: formattedText(record, column.sourceCol);
}

/**
 * Text a spreadsheet would read as a formula gets a leading apostrophe, so pasting rows into Excel
 * or Google Sheets shows the text instead of running it (#208). Formulas start with =, +, -, @, a
 * tab or a carriage return (OWASP, "CSV Injection"). A number's own sign is left alone: when the
 * raw value is a number or Yes/No, "-5" or "-$5.00" is just a negative number.
 */
export function spreadsheetSafe(text: string, raw?: unknown): string {
	if (!/^[=+\-@\t\r]/.test(text)) return text;
	if ((typeof raw === "number" || typeof raw === "boolean") && /^[+-]/.test(text)) return text;
	return `'${text}`;
}

/** Quotes a value for TSV when it contains a tab, line break or quote */
export function tsvField(text: string): string {
	return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/**
 * Columns to copy for a mode. Raw mode leaves out the "(Raw)" twins: their formatted column
 * already copies as the raw value.
 */
export function copyColumnsFor(columns: CopyColumn[], mode: CopyValueMode): CopyColumn[] {
	return mode === "raw" ? columns.filter((c) => !c.id.endsWith("__raw")) : columns;
}

/** Rows as TSV, with a header row */
export function rowsToTsv(
	records: Record<string, unknown>[],
	columns: CopyColumn[],
	mode: CopyValueMode = "shown",
): string {
	const cols = copyColumnsFor(columns, mode);
	const header = cols.map((c) => tsvField(spreadsheetSafe(c.header))).join("\t");
	const lines = records.map((r) => cols.map((c) => tsvField(safeCellText(r, c, mode))).join("\t"));
	return [header, ...lines].join("\n");
}

/** A cell's text for pasting into a spreadsheet (#208) */
const safeCellText = (record: Record<string, unknown>, column: CopyColumn, mode: CopyValueMode): string =>
	spreadsheetSafe(cellText(record, column, mode), record[column.sourceCol]);

/** One column's values, one per line, no header (e.g. to paste into an "in" condition) */
export function columnToText(records: Record<string, unknown>[], column: CopyColumn, mode: CopyValueMode = "shown"): string {
	return records.map((r) => tsvField(safeCellText(r, column, mode))).join("\n");
}

/** "1 row" / "3 rows" */
export const rowCountLabel = (count: number): string => `${count} row${count === 1 ? "" : "s"}`;

/**
 * Writes text to the clipboard: the ToolBox's, or the browser's when the ToolBox's isn't available
 * (host.copyText, #70). Returns whether it worked.
 */
export const writeClipboard = (text: string): Promise<boolean> => copyText(text);
