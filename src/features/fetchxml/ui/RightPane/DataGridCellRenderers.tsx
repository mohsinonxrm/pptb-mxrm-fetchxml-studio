/**
 * Rich cell renderers for different attribute types in DataGrid
 * Renders values using appropriate Fluent UI v9 components
 * Prefers formatted values from OData annotations when available
 */

import { Switch, Badge, Label, SkeletonItem, makeStyles, tokens } from "@fluentui/react-components";
import type { AttributeMetadata } from "../../api/pptbClient";

const useStyles = makeStyles({
	booleanCell: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalXS,
	},
	lookupCell: {
		display: "flex",
		flexDirection: "column",
		gap: "2px",
	},
	lookupName: {
		fontSize: tokens.fontSizeBase200,
		fontWeight: tokens.fontWeightSemibold,
	},
	lookupType: {
		fontSize: tokens.fontSizeBase100,
		color: tokens.colorNeutralForeground3,
	},
	numberCell: {
		textAlign: "right",
		fontVariantNumeric: "tabular-nums",
	},
	dateCell: {
		fontVariantNumeric: "tabular-nums",
	},
	picklistCell: {
		display: "flex",
		alignItems: "center",
	},
});

/**
 * Render a boolean/two-option value as a read-only Switch
 * Uses formatted value from OData annotations if available
 */
export function BooleanCellRenderer({
	value,
	formattedValue,
	attribute,
}: {
	value: unknown;
	formattedValue?: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();
	const boolValue = Boolean(value);

	// Prefer formatted value, otherwise use metadata labels
	const displayValue =
		formattedValue ||
		(boolValue
			? attribute?.OptionSet?.TrueOption?.Label?.UserLocalizedLabel?.Label || "Yes"
			: attribute?.OptionSet?.FalseOption?.Label?.UserLocalizedLabel?.Label || "No");

	return (
		<div className={styles.booleanCell}>
			<Switch checked={boolValue} disabled />
			<Label size="small">{String(displayValue)}</Label>
		</div>
	);
}

/**
 * Render a picklist/state/status value as a Badge
 * Uses formatted value from OData annotations if available
 */
export function PicklistCellRenderer({
	value,
	formattedValue,
	attribute,
}: {
	value: unknown;
	formattedValue?: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();

	// Prefer formatted value
	if (formattedValue) {
		const appearance =
			attribute?.AttributeType === "State" || attribute?.AttributeType === "Status"
				? "filled"
				: "tint";
		return (
			<div className={styles.picklistCell}>
				<Badge appearance={appearance}>{String(formattedValue)}</Badge>
			</div>
		);
	}

	// Fallback to raw value with metadata lookup
	const numValue = typeof value === "number" ? value : parseInt(String(value), 10);

	if (isNaN(numValue)) {
		return <span>—</span>;
	}

	// Find the option label from metadata
	const option = attribute?.OptionSet?.Options?.find((opt) => opt.Value === numValue);
	const label = option?.Label?.UserLocalizedLabel?.Label || String(numValue);

	// Use different badge colors for different states
	const appearance =
		attribute?.AttributeType === "State" || attribute?.AttributeType === "Status"
			? "filled"
			: "tint";

	return (
		<div className={styles.picklistCell}>
			<Badge appearance={appearance}>{label}</Badge>
		</div>
	);
}

/**
 * Render a lookup/customer/owner value
 * Uses formatted value from OData annotations (preferred)
 */
export function LookupCellRenderer({
	value,
	formattedValue,
}: {
	value: unknown;
	formattedValue?: unknown;
}) {
	const styles = useStyles();

	// Prefer formatted value (just the name string)
	if (formattedValue) {
		return <span className={styles.lookupName}>{String(formattedValue)}</span>;
	}

	// Fallback to raw value (structured object)
	if (!value || typeof value !== "object") {
		return <span>—</span>;
	}

	const lookupValue = value as { Id?: string; Name?: string; LogicalName?: string };

	return (
		<div className={styles.lookupCell}>
			<span className={styles.lookupName}>{lookupValue.Name || "—"}</span>
			{lookupValue.LogicalName && (
				<span className={styles.lookupType}>{lookupValue.LogicalName}</span>
			)}
		</div>
	);
}

/**
 * Render a numeric value (integer, decimal, money)
 * Uses formatted value from OData annotations if available
 */
export function NumberCellRenderer({
	value,
	formattedValue,
	attribute,
}: {
	value: unknown;
	formattedValue?: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();

	// Prefer formatted value (already includes currency symbol, etc.)
	if (formattedValue) {
		return <span className={styles.numberCell}>{String(formattedValue)}</span>;
	}

	// Fallback to manual formatting
	if (value === null || value === undefined || value === "") {
		return <span>—</span>;
	}

	const numValue = typeof value === "number" ? value : parseFloat(String(value));

	if (isNaN(numValue)) {
		return <span>—</span>;
	}

	// Format based on attribute type
	if (attribute?.AttributeType === "Money") {
		// Format as currency (would need locale info for proper currency symbol)
		return (
			<span className={styles.numberCell}>
				$
				{numValue.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
			</span>
		);
	}

	if (attribute?.AttributeType === "Integer" || attribute?.AttributeType === "BigInt") {
		return <span className={styles.numberCell}>{numValue.toLocaleString()}</span>;
	}

	// Decimal/Double - use precision from metadata if available
	const precision = attribute?.Precision ?? 2;
	return (
		<span className={styles.numberCell}>
			{numValue.toLocaleString(undefined, {
				minimumFractionDigits: 0,
				maximumFractionDigits: precision,
			})}
		</span>
	);
}

/**
 * Render a date/time value
 * Uses formatted value from OData annotations if available
 */
export function DateTimeCellRenderer({
	value,
	formattedValue,
	attribute,
}: {
	value: unknown;
	formattedValue?: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();

	// Prefer formatted value
	if (formattedValue) {
		return <span className={styles.dateCell}>{String(formattedValue)}</span>;
	}

	// Fallback to manual formatting
	if (!value) {
		return <span>—</span>;
	}

	const dateValue = new Date(String(value));

	if (isNaN(dateValue.getTime())) {
		return <span>—</span>;
	}

	// Check if DateOnly format
	const isDateOnly =
		attribute?.Format === "DateOnly" || attribute?.DateTimeBehavior?.Value === "DateOnly";

	if (isDateOnly) {
		return (
			<span className={styles.dateCell}>
				{dateValue.toLocaleDateString(undefined, {
					year: "numeric",
					month: "short",
					day: "numeric",
				})}
			</span>
		);
	}

	// Full date-time
	return (
		<span className={styles.dateCell}>
			{dateValue.toLocaleString(undefined, {
				year: "numeric",
				month: "short",
				day: "numeric",
				hour: "2-digit",
				minute: "2-digit",
			})}
		</span>
	);
}

/**
 * Default text renderer for strings and other types
 * Uses formatted value from OData annotations if available
 */
export function TextCellRenderer({
	value,
	formattedValue,
}: {
	value: unknown;
	formattedValue?: unknown;
}) {
	// Prefer formatted value
	const displayValue = formattedValue ?? value;

	if (displayValue === null || displayValue === undefined || displayValue === "") {
		return <span>—</span>;
	}

	return <span>{String(displayValue)}</span>;
}

/**
 * Loading skeleton for virtualized cells
 */
export function LoadingCellRenderer() {
	return <SkeletonItem />;
}

/**
 * Get the appropriate cell renderer for an attribute type
 * Now accepts formatted value from OData annotations
 */
export function getCellRenderer(
	attributeType: string | undefined,
	value: unknown,
	formattedValue: unknown | undefined,
	attribute?: AttributeMetadata
) {
	if (value === null || value === undefined) {
		return <span>—</span>;
	}

	switch (attributeType) {
		case "Boolean":
			return (
				<BooleanCellRenderer value={value} formattedValue={formattedValue} attribute={attribute} />
			);

		case "Picklist":
		case "State":
		case "Status":
			return (
				<PicklistCellRenderer value={value} formattedValue={formattedValue} attribute={attribute} />
			);

		case "Lookup":
		case "Customer":
		case "Owner":
			return <LookupCellRenderer value={value} formattedValue={formattedValue} />;

		case "Integer":
		case "BigInt":
		case "Decimal":
		case "Double":
		case "Money":
			return (
				<NumberCellRenderer value={value} formattedValue={formattedValue} attribute={attribute} />
			);

		case "DateTime":
			return (
				<DateTimeCellRenderer value={value} formattedValue={formattedValue} attribute={attribute} />
			);

		default:
			return <TextCellRenderer value={value} formattedValue={formattedValue} />;
	}
}
