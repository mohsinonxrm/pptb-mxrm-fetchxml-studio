/**
 * Local Excel export using exceljs library
 * Exports query results with native Excel data types (numbers, dates, etc.)
 */

import ExcelJS from "exceljs";
import type { AttributeMetadata } from "./pptbClient";
import { getFormattedValue } from "../ui/RightPane/FormattedValueUtils";
import type { DisplaySettings } from "../model/displaySettings";

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
}

/**
 * Export records to Excel with native data types
 * Returns the workbook buffer for download
 */
export async function exportToExcelLocal(
	options: ExportToExcelLocalOptions
): Promise<{ buffer: ArrayBuffer; fileName: string }> {
	const {
		records,
		columns,
		columnDisplayNames,
		attributeMetadata,
		entityName,
		fileName = "export",
		displaySettings,
	} = options;

	// Extract display settings with defaults
	const useLogicalNames = displaySettings?.useLogicalNames ?? false;
	const valueDisplayMode = displaySettings?.valueDisplayMode ?? "formatted";

	const workbook = new ExcelJS.Workbook();
	workbook.creator = "FetchXML Builder";
	workbook.created = new Date();

	const worksheet = workbook.addWorksheet("Data");

	// Helper to get attribute metadata
	const getAttrMeta = (col: string): AttributeMetadata | undefined => {
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

	const exportColumns: ExportColumn[] = [];
	for (const col of columns) {
		// Determine header name based on useLogicalNames setting
		const displayName = columnDisplayNames.get(col) || col;
		const headerName = useLogicalNames ? col : displayName;

		if (valueDisplayMode === "both" && hasFormattedValues(col)) {
			// Add formatted column first, then raw column
			exportColumns.push({ sourceCol: col, headerName, isRawColumn: false });
			exportColumns.push({ sourceCol: col, headerName: `${headerName} (Raw)`, isRawColumn: true });
		} else {
			// Just one column
			exportColumns.push({ sourceCol: col, headerName, isRawColumn: false });
		}
	}

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

	// Helper to get cell value based on display mode and column type
	const getCellValue = (
		record: Record<string, unknown>,
		col: string,
		isRawColumn: boolean,
		attr: AttributeMetadata | undefined
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
					if (rawValue instanceof Date) return rawValue;
					if (typeof rawValue === "string") {
						const date = new Date(rawValue);
						return isNaN(date.getTime()) ? rawValue : date;
					}
					return rawValue;
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
				// Return as Date object for Excel to format properly
				if (rawValue instanceof Date) return rawValue;
				if (typeof rawValue === "string") {
					const date = new Date(rawValue);
					return isNaN(date.getTime()) ? rawValue : date;
				}
				return rawValue;

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

			// Skip formatting for raw columns (show actual values)
			if (ec.isRawColumn) return;

			switch (attr?.AttributeType) {
				case "DateTime":
					cell.numFmt = "yyyy-mm-dd hh:mm:ss";
					break;
				case "Money":
					// Use currency format
					cell.numFmt = '"$"#,##0.00';
					break;
				case "Decimal":
				case "Double": {
					// Get precision from metadata or default to 2
					const precision = (attr as { Precision?: number })?.Precision ?? 2;
					cell.numFmt = `0.${"0".repeat(precision)}`;
					break;
				}
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
export function downloadExcelFile(buffer: ArrayBuffer, fileName: string): void {
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
