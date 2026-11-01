/**
 * Rich cell renderers for different attribute types in DataGrid
 * Renders values using appropriate Fluent UI v9 components
 * Prefers formatted values from OData annotations when available
 */

import { Switch, Badge, Label, makeStyles, tokens } from "@fluentui/react-components";
import type { AttributeMetadata } from "../../api/pptbClient";
import { HighlightText } from "./HighlightText";
import { useGridDensity } from "./gridDensityContext";

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
		minWidth: 0, // required for flex child to clip overflow
	},
	lookupName: {
		display: "block",
		fontSize: tokens.fontSizeBase200,
		fontWeight: tokens.fontWeightSemibold,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	lookupType: {
		display: "block",
		fontSize: tokens.fontSizeBase100,
		color: tokens.colorNeutralForeground3,
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	numberCell: {
		display: "block",
		textAlign: "right",
		fontVariantNumeric: "tabular-nums",
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	dateCell: {
		display: "block",
		fontVariantNumeric: "tabular-nums",
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
	},
	picklistCell: {
		display: "flex",
		alignItems: "center",
	},
	truncateText: {
		display: "block",
		overflow: "hidden",
		textOverflow: "ellipsis",
		whiteSpace: "nowrap",
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
	const density = useGridDensity();
	const boolValue = Boolean(value);

	// Prefer formatted value, otherwise use metadata labels
	const displayValue =
		formattedValue ||
		(boolValue
			? attribute?.OptionSet?.TrueOption?.Label?.UserLocalizedLabel?.Label || "Yes"
			: attribute?.OptionSet?.FalseOption?.Label?.UserLocalizedLabel?.Label || "No");

	return (
		<div className={styles.booleanCell}>
			{/* The small switch fits the Small and Extra small rows (#228) */}
			<Switch checked={boolValue} disabled size={density === "medium" ? "medium" : "small"} />
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
	// The small badge fits the Extra small rows (#228)
	const badgeSize = useGridDensity() === "extra-small" ? "small" : "medium";

	// Prefer formatted value
	if (formattedValue) {
		const appearance =
			attribute?.AttributeType === "State" || attribute?.AttributeType === "Status"
				? "filled"
				: "tint";
		return (
			<div className={styles.picklistCell}>
				<Badge appearance={appearance} size={badgeSize}>{String(formattedValue)}</Badge>
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
			<Badge appearance={appearance} size={badgeSize}>{label}</Badge>
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
	const density = useGridDensity();

	// Prefer formatted value (just the name string)
	if (formattedValue) {
		return <span className={styles.lookupName}><HighlightText text={String(formattedValue)} /></span>;
	}

	// Fallback to raw value (structured object)
	if (!value || typeof value !== "object") {
		return <span>—</span>;
	}

	const lookupValue = value as { Id?: string; Name?: string; LogicalName?: string };

	return (
		<div className={styles.lookupCell}>
			<span className={styles.lookupName}><HighlightText text={lookupValue.Name || "—"} /></span>
			{/* A second line doesn't fit an Extra small row (#228) */}
			{lookupValue.LogicalName && density !== "extra-small" && (
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
		return <span className={styles.numberCell}><HighlightText text={String(formattedValue)} /></span>;
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
		// No formatted value, so no currency: the amount alone, not a "$" (#184)
		return (
			<span className={styles.numberCell}>
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
		return <span className={styles.dateCell}><HighlightText text={String(formattedValue)} /></span>;
	}

	// Fallback to manual formatting
	if (!value) {
		return <span>—</span>;
	}

	// A Date Only value ("2024-03-14") is that date: new Date() would read it as UTC midnight and
	// show the day before west of UTC (#184, as #198)
	const dateOnlyText = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value));
	const dateValue = dateOnlyText
		? new Date(Number(dateOnlyText[1]), Number(dateOnlyText[2]) - 1, Number(dateOnlyText[3]))
		: new Date(String(value));

	if (isNaN(dateValue.getTime())) {
		return <span>—</span>;
	}

	// Check if DateOnly format
	const isDateOnly =
		!!dateOnlyText || attribute?.Format === "DateOnly" || attribute?.DateTimeBehavior?.Value === "DateOnly";

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
	const styles = useStyles();
	// Prefer formatted value
	const displayValue = formattedValue ?? value;

	if (displayValue === null || displayValue === undefined || displayValue === "") {
		return <span>—</span>;
	}

	return <span className={styles.truncateText}><HighlightText text={String(displayValue)} /></span>;
}

