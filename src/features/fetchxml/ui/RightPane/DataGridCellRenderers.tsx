/**
 * Rich cell renderers for different attribute types in DataGrid
 * Renders values using appropriate Fluent UI v9 components
 */

import { Switch, Badge, Label, makeStyles, tokens } from "@fluentui/react-components";
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
 */
export function BooleanCellRenderer({
	value,
	attribute,
}: {
	value: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();
	const boolValue = Boolean(value);

	// Get labels from metadata if available
	const trueLabel = attribute?.OptionSet?.TrueOption?.Label?.UserLocalizedLabel?.Label || "Yes";
	const falseLabel = attribute?.OptionSet?.FalseOption?.Label?.UserLocalizedLabel?.Label || "No";

	return (
		<div className={styles.booleanCell}>
			<Switch checked={boolValue} disabled />
			<Label size="small">{boolValue ? trueLabel : falseLabel}</Label>
		</div>
	);
}

/**
 * Render a picklist/state/status value as a Badge
 */
export function PicklistCellRenderer({
	value,
	attribute,
}: {
	value: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();
	const numValue = typeof value === "number" ? value : parseInt(String(value), 10);

	if (isNaN(numValue)) {
		return <span>—</span>;
	}

	// Find the option label from metadata
	const option = attribute?.OptionSet?.Options?.find((opt) => opt.Value === numValue);
	const label = option?.Label?.UserLocalizedLabel?.Label || String(numValue);

	// Use different badge colors for different states
	const appearance =
		attribute?.AttributeType === "State" ? (numValue === 0 ? "filled" : "outline") : "tint";

	return (
		<div className={styles.picklistCell}>
			<Badge appearance={appearance}>{label}</Badge>
		</div>
	);
}

/**
 * Render a lookup/customer/owner value
 */
export function LookupCellRenderer({ value }: { value: unknown }) {
	const styles = useStyles();

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
 */
export function NumberCellRenderer({
	value,
	attribute,
}: {
	value: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();

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
 */
export function DateTimeCellRenderer({
	value,
	attribute,
}: {
	value: unknown;
	attribute?: AttributeMetadata;
}) {
	const styles = useStyles();

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
 */
export function TextCellRenderer({ value }: { value: unknown }) {
	if (value === null || value === undefined || value === "") {
		return <span>—</span>;
	}

	return <span>{String(value)}</span>;
}

/**
 * Get the appropriate cell renderer for an attribute type
 */
export function getCellRenderer(
	attributeType: string | undefined,
	value: unknown,
	attribute?: AttributeMetadata
) {
	if (value === null || value === undefined) {
		return <span>—</span>;
	}

	switch (attributeType) {
		case "Boolean":
			return <BooleanCellRenderer value={value} attribute={attribute} />;

		case "Picklist":
		case "State":
		case "Status":
			return <PicklistCellRenderer value={value} attribute={attribute} />;

		case "Lookup":
		case "Customer":
		case "Owner":
			return <LookupCellRenderer value={value} />;

		case "Integer":
		case "BigInt":
		case "Decimal":
		case "Double":
		case "Money":
			return <NumberCellRenderer value={value} attribute={attribute} />;

		case "DateTime":
			return <DateTimeCellRenderer value={value} attribute={attribute} />;

		default:
			return <TextCellRenderer value={value} />;
	}
}
