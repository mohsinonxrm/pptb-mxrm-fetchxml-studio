/**
 * Utilities for extracting formatted values and column display names from OData annotated records
 *
 * OData Formatted Value Annotations:
 * - Regular columns: {column}@OData.Community.Display.V1.FormattedValue
 * - Lookup columns: _{column}_value@OData.Community.Display.V1.FormattedValue
 * - Aliased columns: {alias}@OData.Community.Display.V1.FormattedValue
 * - Attribute name: {column/alias}@OData.Community.Display.V1.AttributeName
 *
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/fetchxml/formatted-values
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/fetchxml/column-comparison-aliases
 */

import type { AttributeMetadata } from "../../api/pptbClient";

const FORMATTED_VALUE_SUFFIX = "@OData.Community.Display.V1.FormattedValue";
const ATTRIBUTE_NAME_SUFFIX = "@OData.Community.Display.V1.AttributeName";

/**
 * Extracts the formatted value for a column from an OData annotated record.
 * Handles regular columns, lookup columns, and aliased columns.
 * Falls back to raw value if formatted value is not available.
 *
 * @param record - The OData record with annotations
 * @param column - The column name (logical name or alias)
 * @returns The formatted value if available, otherwise the raw value
 *
 * @example
 * // Regular column
 * getFormattedValue({ statuscode: 1, "statuscode@...FormattedValue": "Active" }, "statuscode")
 * // Returns: "Active"
 *
 * @example
 * // Lookup column
 * getFormattedValue({ _ownerid_value: "guid", "_ownerid_value@...FormattedValue": "John Doe" }, "_ownerid_value")
 * // Returns: "John Doe"
 *
 * @example
 * // Aliased column
 * getFormattedValue({ owner_name: "guid", "owner_name@...FormattedValue": "John Doe" }, "owner_name")
 * // Returns: "John Doe"
 */
export function getFormattedValue(record: Record<string, unknown>, column: string): unknown {
	if (!record || !column) return null;

	// Try formatted value with exact column name
	const formattedKey = `${column}${FORMATTED_VALUE_SUFFIX}`;
	if (formattedKey in record) {
		return record[formattedKey];
	}

	// For lookup columns that might be in _{column}_value format
	// Check if this is already a lookup pattern or try adding it
	if (column.startsWith("_") && column.endsWith("_value")) {
		// Already in lookup format, formatted key already checked above
		return record[column]; // Return raw value
	}

	// Try adding lookup wrapper
	const lookupFormattedKey = `_${column}_value${FORMATTED_VALUE_SUFFIX}`;
	if (lookupFormattedKey in record) {
		return record[lookupFormattedKey];
	}

	// Fall back to raw value
	return record[column];
}

/**
 * Determines the display name for a column header.
 * Tries in order:
 * 1. AttributeName annotation from the record (for aliases)
 * 2. DisplayName from AttributeMetadata
 * 3. The raw column name
 *
 * @param column - The column name (logical name or alias)
 * @param attributeMetadata - Optional attribute metadata map
 * @param record - Optional record with OData annotations (for AttributeName)
 * @returns The display name for the column
 *
 * @example
 * // With AttributeName annotation (aliased column from link-entity)
 * getColumnDisplayName("owner_name", metadata, { "owner_name@...AttributeName": "ownerid" })
 * // Returns metadata display name for "ownerid"
 *
 * @example
 * // With metadata only
 * getColumnDisplayName("statuscode", metadata)
 * // Returns: "Status Reason" (from metadata)
 *
 * @example
 * // Fallback to column name
 * getColumnDisplayName("customfield", undefined, undefined)
 * // Returns: "customfield"
 */
export function getColumnDisplayName(
	column: string,
	attributeMetadata?: Map<string, AttributeMetadata>,
	record?: Record<string, unknown>
): string {
	if (!column) return "";

	// Try to get the original attribute name from OData annotation (for aliased columns)
	let originalAttributeName = column;
	if (record) {
		const attributeNameKey = `${column}${ATTRIBUTE_NAME_SUFFIX}`;
		if (attributeNameKey in record && typeof record[attributeNameKey] === "string") {
			originalAttributeName = record[attributeNameKey] as string;
		}
	}

	// Try to get display name from metadata
	if (attributeMetadata) {
		// First try the original attribute name (in case it's an alias)
		const metadataByOriginal = attributeMetadata.get(originalAttributeName);
		if (metadataByOriginal?.DisplayName?.UserLocalizedLabel?.Label) {
			return metadataByOriginal.DisplayName.UserLocalizedLabel.Label;
		}

		// Then try the column name directly
		const metadataByColumn = attributeMetadata.get(column);
		if (metadataByColumn?.DisplayName?.UserLocalizedLabel?.Label) {
			return metadataByColumn.DisplayName.UserLocalizedLabel.Label;
		}

		// For lookup columns in _{name}_value format, try extracting the base name
		if (column.startsWith("_") && column.endsWith("_value")) {
			const baseName = column.slice(1, -6); // Remove _ prefix and _value suffix
			const metadataByBase = attributeMetadata.get(baseName);
			if (metadataByBase?.DisplayName?.UserLocalizedLabel?.Label) {
				return metadataByBase.DisplayName.UserLocalizedLabel.Label;
			}
		}
	}

	// Fall back to the column name
	// For lookup columns, show a cleaner version
	if (column.startsWith("_") && column.endsWith("_value")) {
		return column.slice(1, -6); // Remove _ prefix and _value suffix
	}

	return column;
}

/**
 * Helper to check if a record has any formatted values available.
 * Useful for optimization - if no formatted values exist, can skip formatting logic.
 *
 * @param record - The OData record to check
 * @returns true if the record has at least one formatted value annotation
 */
export function hasFormattedValues(record: Record<string, unknown>): boolean {
	if (!record) return false;
	return Object.keys(record).some((key) => key.endsWith(FORMATTED_VALUE_SUFFIX));
}

/**
 * Checks if a column name is an OData or CRM annotation that should be hidden from display.
 * Filters out:
 * - @OData.Community.Display.V1.FormattedValue
 * - @OData.Community.Display.V1.AttributeName
 * - @Microsoft.Dynamics.CRM.associatednavigationproperty
 * - @Microsoft.Dynamics.CRM.lookuplogicalname
 * - etc.
 *
 * @param columnName - The column name to check
 * @returns true if this is an annotation column that should be hidden
 */
export function isAnnotationColumn(columnName: string): boolean {
	if (!columnName) return false;
	return (
		columnName.includes("@OData.") ||
		columnName.includes("@Microsoft.Dynamics.CRM.") ||
		columnName.includes("@odata.") ||
		columnName.includes("@microsoft.dynamics.crm.")
	);
}

/**
 * Filters out annotation columns from a list of column names.
 * Returns only the actual data columns that should be displayed in the grid.
 *
 * @param columns - Array of column names from query results
 * @returns Filtered array with only displayable columns
 */
export function filterDisplayableColumns(columns: string[]): string[] {
	if (!columns) return [];
	return columns.filter((col) => !isAnnotationColumn(col));
}
