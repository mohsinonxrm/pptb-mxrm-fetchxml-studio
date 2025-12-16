/**
 * Input picker for operators that require TWO values (e.g., between, not-between, in-fiscal-period-and-year)
 * Displays two input fields side-by-side with an "and" label between them
 */

import { Label, makeStyles, tokens } from "@fluentui/react-components";
import { NumericInputPicker } from "./NumericInputPicker";
import { DateTimeInputPicker } from "./DateTimeInputPicker";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";

const useStyles = makeStyles({
	container: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalM,
		width: "100%",
	},
	inputWrapper: {
		flex: 1,
		minWidth: "120px",
	},
	separator: {
		color: tokens.colorNeutralForeground3,
		fontWeight: tokens.fontWeightSemibold,
		paddingInline: tokens.spacingHorizontalXS,
	},
});

interface BetweenInputPickerProps {
	attribute: AttributeMetadata;
	value?: [number | string | undefined, number | string | undefined]; // Tuple [value1, value2]
	onChange: (value: [number | string | undefined, number | string | undefined]) => void;
	placeholder1?: string;
	placeholder2?: string;
}

/**
 * BetweenInputPicker component for dual-value operators like 'between', 'not-between'
 */
export function BetweenInputPicker({
	attribute,
	value = [undefined, undefined],
	onChange,
	placeholder1 = "Start value",
	placeholder2 = "End value",
}: BetweenInputPickerProps) {
	const styles = useStyles();

	const [value1, value2] = value;

	const handleValue1Change = (newValue: number | string | Date | undefined) => {
		const val1 =
			newValue instanceof Date
				? newValue.toISOString()
				: typeof newValue === "number" || typeof newValue === "string"
				? newValue
				: undefined;
		onChange([val1, value2]);
	};

	const handleValue2Change = (newValue: number | string | Date | undefined) => {
		const val2 =
			newValue instanceof Date
				? newValue.toISOString()
				: typeof newValue === "number" || typeof newValue === "string"
				? newValue
				: undefined;
		onChange([value1, val2]);
	};

	const isNumeric =
		attribute.AttributeType === "Integer" ||
		attribute.AttributeType === "BigInt" ||
		attribute.AttributeType === "Decimal" ||
		attribute.AttributeType === "Double";

	const isDateTime = attribute.AttributeType === "DateTime";

	return (
		<div className={styles.container}>
			<div className={styles.inputWrapper}>
				{isNumeric ? (
					<NumericInputPicker
						attribute={attribute}
						value={typeof value1 === "number" ? value1 : undefined}
						onChange={handleValue1Change}
						placeholder={placeholder1}
					/>
				) : isDateTime ? (
					<DateTimeInputPicker
						attribute={attribute}
						value={typeof value1 === "string" ? value1 : undefined}
						onChange={handleValue1Change}
						placeholder={placeholder1}
					/>
				) : null}
			</div>

			<Label className={styles.separator}>and</Label>

			<div className={styles.inputWrapper}>
				{isNumeric ? (
					<NumericInputPicker
						attribute={attribute}
						value={typeof value2 === "number" ? value2 : undefined}
						onChange={handleValue2Change}
						placeholder={placeholder2}
					/>
				) : isDateTime ? (
					<DateTimeInputPicker
						attribute={attribute}
						value={typeof value2 === "string" ? value2 : undefined}
						onChange={handleValue2Change}
						placeholder={placeholder2}
					/>
				) : null}
			</div>
		</div>
	);
}
