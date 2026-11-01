/**
 * Local Excel export using exceljs library
 * Exports query results with native Excel data types (numbers, dates, etc.)
 */

import ExcelJS from "exceljs";
import type { AttributeMetadata, FormattingMetadata } from "./pptbClient";
import { decimalNumberFormat, moneyNumberFormat } from "./excelFormats";
import {
	UTC_TIME_ZONE,
	dateTimeBehavior,
	resultWallClock,
	wallClockToExcelDate,
	type DataverseTimeZone,
} from "../model/dateTimeValues";
import { getFormattedValue } from "../ui/RightPane/FormattedValueUtils";
import type { DisplaySettings } from "../model/displaySettings";
import { debugLog } from "../../../shared/utils/debug";

interface ExportToExcelLocalOptions {
	/** Records to export */
	records: Record<string, unknown>[];
	/** Column keys in the order they should appear */
	columns: string[];
	/** Map of column key to display name for header */
	columnDisplayNames: Map<string, string>;
	/** Multi-entity attribute metadata for type conversion */
	attributeMetadata?: Map<string, Map<string, AttributeMetadata>>;
	/** Entity logical name for single-entity metadata lookup */
	entityName?: string;
	/** Filename without extension */
	fileName?: string;
	/** Display settings for column headers and value modes */
	displaySettings?: DisplaySettings;
	/** Date formats and behaviours, money precision, per entity (loadFormattingMetadata, #184) */
	formattingMetadata?: Map<string, Map<string, FormattingMetadata>>;
	/** The user's Dataverse time zone, for User Local date-times (#184); UTC when unknown */
	timeZone?: DataverseTimeZone;
}

/**
 * Export records to Excel with native data types
 * Returns the workbook buffer for download
 */
export async function exportToExcelLocal(
	options: ExportToExcelLocalOptions,
): Promise<{ buffer: ArrayBuffer; fileName: string }> {
	const {
		records,
		columns,
		columnDisplayNames,
		attributeMetadata,
		entityName,
		fileName = "export",
		displaySettings,
		formattingMetadata,
		timeZone,
	} = options;

	// Extract display settings with defaults
	const useLogicalNames = displaySettings?.useLogicalNames ?? false;
	const valueDisplayMode = displaySettings?.valueDisplayMode ?? "formatted";

	const workbook = new ExcelJS.Workbook();
	workbook.creator = "FetchXML Studio";
	workbook.created = new Date();

	const worksheet = workbook.addWorksheet("Data");

	// An attribute's metadata, with its date and money formatting (#184)
	const getAttrMeta = (col: string): (AttributeMetadata & FormattingMetadata) | undefined => {
		const attr = findAttrMeta(col);
		if (!attr) return undefined;
		const formatting = findFormatting(attr.LogicalName);
		return formatting ? { ...attr, ...formatting } : attr;
	};
	const findFormatting = (logicalName: string): FormattingMetadata | undefined => {
		if (!formattingMetadata) return undefined;
		const own = entityName ? formattingMetadata.get(entityName)?.get(logicalName) : undefined;
		if (own) return own;
		for (const [, entityFormatting] of formattingMetadata) {
			const found = entityFormatting.get(logicalName);
			if (found) return found;
		}
		return undefined;
	};

	// Helper to get attribute metadata
	const findAttrMeta = (col: string): AttributeMetadata | undefined => {
		if (!attributeMetadata) return undefined;

		// Try with entity name first
		if (entityName) {
			const entityAttrs = attributeMetadata.get(entityName);
			if (entityAttrs) {
				// Direct column name
				if (entityAttrs.has(col)) return entityAttrs.get(col);
				// Lookup field (_attr_value -> attr)
				if (col.startsWith("_") && col.endsWith("_value")) {
					const baseAttr = col.slice(1, -6);
					if (entityAttrs.has(baseAttr)) return entityAttrs.get(baseAttr);
				}
			}
		}

		// Try to find in any entity's metadata
		for (const [, entityAttrs] of attributeMetadata) {
			if (entityAttrs.has(col)) return entityAttrs.get(col);
			if (col.startsWith("_") && col.endsWith("_value")) {
				const baseAttr = col.slice(1, -6);
				if (entityAttrs.has(baseAttr)) return entityAttrs.get(baseAttr);
			}
		}

		return undefined;
	};

	// Helper to check if a column has formatted values in the data
	const hasFormattedValues = (col: string): boolean => {
		return records.some((record) => {
			const rawValue = record[col];
			const formattedValue = getFormattedValue(record, col);
			return formattedValue !== undefined && formattedValue !== rawValue;
		});
	};

	// Build export columns list - for "both" mode, add raw columns where formatted values exist
	interface ExportColumn {
		sourceCol: string;
		headerName: string;
		isRawColumn: boolean;
	}

	// Raw values are UTC for User Local date-times; formatted ones are in the user's Dataverse time
	// zone, or UTC when it couldn't be read (#184)
	const zone = timeZone ?? UTC_TIME_ZONE;
	const isUtcDateTime = (col: string, raw: boolean): boolean => {
		const attr = getAttrMeta(col);
		return (
			attr?.AttributeType === "DateTime" &&
			dateTimeBehavior(attr) === "UserLocal" &&
			attr.Format !== "DateOnly" &&
			(raw || !timeZone)
		);
	};
	const withUtc = (header: string, col: string, raw: boolean) => (isUtcDateTime(col, raw) ? `${header} (UTC)` : header);

	const exportColumns: ExportColumn[] = [];
	for (const col of columns) {
		// Determine header name based on useLogicalNames setting
		const displayName = columnDisplayNames.get(col) || col;
		const headerName = useLogicalNames ? col : displayName;

		if (valueDisplayMode === "both" && hasFormattedValues(col)) {
			// Add formatted column first, then raw column
			exportColumns.push({ sourceCol: col, headerName: withUtc(headerName, col, false), isRawColumn: false });
			exportColumns.push({ sourceCol: col, headerName: withUtc(`${headerName} (Raw)`, col, true), isRawColumn: true });
		} else {
			// Just one column; in raw mode it holds raw values
			exportColumns.push({ sourceCol: col, headerName: withUtc(headerName, col, valueDisplayMode === "raw"), isRawColumn: false });
		}
	}
	/** Whether a column holds raw values: a (Raw) twin, or every column in raw mode */
	const showsRaw = (ec: ExportColumn) => ec.isRawColumn || valueDisplayMode === "raw";

	// Create header row
	const headerRow = worksheet.addRow(exportColumns.map((ec) => ec.headerName));
	headerRow.font = { bold: true };
	headerRow.fill = {
		type: "pattern",
		pattern: "solid",
		fgColor: { argb: "FFE0E0E0" },
	};

	// Set column widths based on attribute type and header length
	worksheet.columns = exportColumns.map((ec, index) => {
		const attr = getAttrMeta(ec.sourceCol);
		let width = Math.max(ec.headerName.length + 2, 10);

		// Adjust width based on attribute type
		switch (attr?.AttributeType) {
			case "DateTime":
				width = Math.max(width, 18);
				break;
			case "Money":
			case "Decimal":
			case "Double":
				width = Math.max(width, 14);
				break;
			case "Lookup":
			case "Customer":
			case "Owner":
				width = Math.max(width, 25);
				break;
			case "Memo":
				width = Math.max(width, 40);
				break;
		}

		return { key: `col_${index}`, width };
	});

	/**
	 * A date cell (#184). exceljs writes a Date's UTC reading, so the Date carries the wall clock to
	 * show: the date for Date Only values and formats, the stored reading for Time Zone Independent
	 * ones, and for User Local date-times the user's Dataverse time zone (UTC when raw).
	 */
	const excelDate = (
		rawValue: unknown,
		attr: (AttributeMetadata & FormattingMetadata) | undefined,
		raw: boolean,
	): unknown => {
		if (rawValue instanceof Date) return rawValue;
		if (typeof rawValue !== "string") return rawValue;
		const behavior = dateTimeBehavior(attr);
		const dateFormat = attr?.Format === "DateOnly";
		const wall = resultWallClock(rawValue, behavior, raw && behavior === "UserLocal" && !dateFormat ? UTC_TIME_ZONE : zone);
		if (!wall) return rawValue;
		return wallClockToExcelDate(dateFormat ? { ...wall, hour: 0, minute: 0, second: 0 } : wall);
	};

	// Helper to get cell value based on display mode and column type
	const getCellValue = (
		record: Record<string, unknown>,
		col: string,
		isRawColumn: boolean,
		attr: (AttributeMetadata & FormattingMetadata) | undefined,
	): unknown => {
		const rawValue = record[col];
		const formattedValue = getFormattedValue(record, col);

		// Handle null/undefined
		if (rawValue === null || rawValue === undefined) {
			return "";
		}

		// For raw columns or raw display mode, return raw value with type conversion
		if (isRawColumn || valueDisplayMode === "raw") {
			switch (attr?.AttributeType) {
				case "Integer":
				case "BigInt":
					return typeof rawValue === "number" ? rawValue : parseInt(String(rawValue), 10);
				case "Decimal":
				case "Double":
				case "Money":
					return typeof rawValue === "number" ? rawValue : parseFloat(String(rawValue));
				case "DateTime":
					return excelDate(rawValue, attr, true);
				case "Boolean":
					return rawValue;
				default:
					return rawValue;
			}
		}

		// For formatted or "both" mode (main column), return formatted values
		switch (attr?.AttributeType) {
			case "Boolean":
				return formattedValue || (rawValue ? "Yes" : "No");

			case "Integer":
			case "BigInt":
				// Numbers still get typed, but formatted value is used if different
				return typeof rawValue === "number" ? rawValue : parseInt(String(rawValue), 10);

			case "Decimal":
			case "Double":
			case "Money":
				return typeof rawValue === "number" ? rawValue : parseFloat(String(rawValue));

			case "DateTime":
				return excelDate(rawValue, attr, false);

			case "Picklist":
			case "State":
			case "Status":
				// Return formatted value (label) for choice fields
				return formattedValue || rawValue;

			case "Lookup":
			case "Customer":
			case "Owner":
				// Return formatted value (name) for lookup fields
				return formattedValue || rawValue;

			default:
				// For strings and unknown types
				return formattedValue !== undefined && formattedValue !== rawValue
					? formattedValue
					: String(rawValue);
		}
	};

	// Add data rows with proper types
	for (const record of records) {
		const rowValues: unknown[] = exportColumns.map((ec) => {
			const attr = getAttrMeta(ec.sourceCol);
			return getCellValue(record, ec.sourceCol, ec.isRawColumn, attr);
		});

		const dataRow = worksheet.addRow(rowValues);

		// Apply number format based on attribute type (only for non-raw columns in formatted mode)
		exportColumns.forEach((ec, index) => {
			const attr = getAttrMeta(ec.sourceCol);
			const cell = dataRow.getCell(index + 1);

			// Dates always need a format, or Excel shows a serial number
			if (attr?.AttributeType === "DateTime" && cell.value instanceof Date) {
				const dateOnly = dateTimeBehavior(attr) === "DateOnly" || attr.Format === "DateOnly";
				cell.numFmt = dateOnly ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss";
				return;
			}

			// Raw values aren't formatted (raw mode, and the (Raw) columns)
			if (showsRaw(ec)) return;

			switch (attr?.AttributeType) {
				case "Money": {
					// The record's own currency symbol and decimals, from the formatted value (#184)
					const raw = record[ec.sourceCol];
					if (typeof raw !== "number") break;
					const precision = attr.PrecisionSource === 0 ? attr.Precision : undefined;
					cell.numFmt = moneyNumberFormat(getFormattedValue(record, ec.sourceCol), raw, precision);
					break;
				}
				case "Decimal":
				case "Double":
					// The column's precision; "0" when it has none (#184)
					cell.numFmt = decimalNumberFormat(attr.Precision);
					break;
				case "Integer":
				case "BigInt":
					cell.numFmt = "#,##0";
					break;
			}
		});
	}

	// Freeze the header row
	worksheet.views = [{ state: "frozen", ySplit: 1 }];

	// Auto-filter
	if (exportColumns.length > 0) {
		worksheet.autoFilter = {
			from: { row: 1, column: 1 },
			to: { row: records.length + 1, column: exportColumns.length },
		};
	}

	// Generate buffer
	const buffer = await workbook.xlsx.writeBuffer();
	// Include full timestamp (date + time) to avoid overwrite prompts on repeated exports
	const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
	const finalFileName = `${fileName}_${timestamp}.xlsx`;

	return {
		buffer: buffer as ArrayBuffer,
		fileName: finalFileName,
	};
}

/**
 * Trigger browser download of the Excel file
 */
export async function downloadExcelFile(buffer: ArrayBuffer, fileName: string): Promise<void> {
	// Use PPTB fileSystem.saveFile when running in the desktop host (opens native save dialog)
	if (window.toolboxAPI?.fileSystem?.saveFile) {
		const saved = await window.toolboxAPI.fileSystem.saveFile(fileName, new Uint8Array(buffer));
		if (saved) {
			debugLog("exportAPI", `✅ File saved via fileSystem: ${saved}`);
		}
		return;
	}

	// Fallback: browser blob download (standalone / dev context)
	const blob = new Blob([buffer], {
		type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	});
	const url = URL.createObjectURL(blob);

	const link = document.createElement("a");
	link.href = url;
	link.download = fileName;
	document.body.appendChild(link);
	link.click();

	// Cleanup
	document.body.removeChild(link);
	URL.revokeObjectURL(url);
}
