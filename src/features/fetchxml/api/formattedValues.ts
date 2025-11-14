/**
 * Utilities for working with Dataverse formatted values
 * Formatted values come from Web API with @OData.Community.Display.V1.FormattedValue annotations
 */

/**
 * Get the formatted value for a property if available
 * Falls back to raw value if no formatted value exists
 * Handles both regular properties and lookup properties (with _propertyname_value pattern)
 * 
 * @example
 * // Record from API:
 * // {
 * //   "statuscode": 1,
 * //   "statuscode@OData.Community.Display.V1.FormattedValue": "Active",
 * //   "createdon": "2024-01-15T10:30:00Z",
 * //   "createdon@OData.Community.Display.V1.FormattedValue": "1/15/2024 10:30 AM",
 * //   "_createdby_value": "guid-here",
 * //   "_createdby_value@OData.Community.Display.V1.FormattedValue": "John Doe"
 * // }
 * 
 * getFormattedValue(record, "statuscode") // Returns "Active"
 * getFormattedValue(record, "createdon") // Returns "1/15/2024 10:30 AM"
 * getFormattedValue(record, "_createdby_value") // Returns "John Doe"
 * getFormattedValue(record, "name") // Returns raw name value
 */
export function getFormattedValue(
	record: Record<string, unknown>,
	propertyName: string
): unknown {
	const formattedKey = `${propertyName}@OData.Community.Display.V1.FormattedValue`;
	
	// Return formatted value if it exists, otherwise return raw value
	if (formattedKey in record && record[formattedKey] !== null && record[formattedKey] !== undefined) {
		return record[formattedKey];
	}
	
	return record[propertyName];
}

/**
 * Check if a property has a formatted value annotation
 */
export function hasFormattedValue(
	record: Record<string, unknown>,
	propertyName: string
): boolean {
	const formattedKey = `${propertyName}@OData.Community.Display.V1.FormattedValue`;
	return formattedKey in record;
}

/**
 * Get both raw and formatted values for a property
 */
export function getValuePair(
	record: Record<string, unknown>,
	propertyName: string
): {
	raw: unknown;
	formatted?: unknown;
} {
	const raw = record[propertyName];
	const formattedKey = `${propertyName}@OData.Community.Display.V1.FormattedValue`;
	const formatted = record[formattedKey];
	
	return {
		raw,
		formatted: formatted !== null && formatted !== undefined ? formatted : undefined,
	};
}

/**
 * Transform a record to use formatted values where available
 * Creates a new object with formatted values replacing raw values
 * Useful for displaying data in grids
 */
export function transformRecordWithFormattedValues(
	record: Record<string, unknown>
): Record<string, unknown> {
	const transformed: Record<string, unknown> = {};
	
	// Get all property names (excluding OData annotations)
	const propertyNames = Object.keys(record).filter(
		(key) => !key.includes("@OData") && !key.includes("@Microsoft")
	);
	
	for (const prop of propertyNames) {
		transformed[prop] = getFormattedValue(record, prop);
	}
	
	return transformed;
}

/**
 * Check if the response includes formatted values
 * Useful for detecting if the API was called with Prefer: odata.include-annotations="*"
 */
export function responseHasFormattedValues(
	records: Record<string, unknown>[]
): boolean {
	if (records.length === 0) return false;
	
	const firstRecord = records[0];
	const keys = Object.keys(firstRecord);
	
	return keys.some((key) => key.includes("@OData.Community.Display.V1.FormattedValue"));
}
