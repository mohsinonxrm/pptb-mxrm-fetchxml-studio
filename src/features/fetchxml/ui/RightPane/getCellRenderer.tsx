import type { AttributeMetadata } from "../../api/pptbClient";
import {
	BooleanCellRenderer,
	DateTimeCellRenderer,
	LookupCellRenderer,
	NumberCellRenderer,
	PicklistCellRenderer,
	TextCellRenderer,
} from "./DataGridCellRenderers";

/**
 * Get the appropriate cell renderer for an attribute type
 * Now accepts formatted value from OData annotations
 */
export function getCellRenderer(
	attributeType: string | undefined,
	value: unknown,
	formattedValue: unknown | undefined,
	attribute?: AttributeMetadata,
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
