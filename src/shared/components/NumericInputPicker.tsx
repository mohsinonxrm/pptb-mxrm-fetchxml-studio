/**
 * Shared Numeric Input Picker component using Fluent UI v9 Input
 * Handles Integer, BigInt, Decimal, and Double attribute types
 * Sets min/max values and precision based on metadata
 */

import { useState, useEffect } from "react";
import { Input, useId } from "@fluentui/react-components";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";

interface NumericInputPickerProps {
	attribute: AttributeMetadata;
	value?: number | string;
	onChange: (value: number | undefined) => void;
	placeholder?: string;
	disabled?: boolean;
}

export function NumericInputPicker({
	attribute,
	value,
	onChange,
	placeholder = "Enter a number",
	disabled = false,
}: NumericInputPickerProps) {
	const inputId = useId("numeric-input");
	const [inputValue, setInputValue] = useState<string>("");

	// Format number to avoid scientific notation
	const formatNumber = (num: number): string => {
		if (Math.abs(num) < 1e-10 && num !== 0) {
			// Very small numbers - use high precision
			return num.toFixed(20).replace(/\.?0+$/, "");
		} else if (Number.isInteger(num)) {
			// Integer - no decimal point needed
			return num.toString();
		} else {
			// Decimal - preserve precision but remove trailing zeros
			return num.toFixed(10).replace(/\.?0+$/, "");
		}
	};

	// Sync input value with prop value
	useEffect(() => {
		if (value !== undefined && value !== null && value !== "") {
			const numValue = typeof value === "number" ? value : parseFloat(String(value));
			setInputValue(isNaN(numValue) ? String(value) : formatNumber(numValue));
		} else {
			setInputValue("");
		}
	}, [value]);

	// Determine input step based on attribute type and precision
	const getStep = (): number | undefined => {
		const attrType = attribute.AttributeType;

		// Integer and BigInt: whole numbers only
		if (attrType === "Integer" || attrType === "BigInt") {
			return 1;
		}

		// Decimal and Double: use precision if available
		if ((attrType === "Decimal" || attrType === "Double") && attribute.Precision !== undefined) {
			// step = 10^(-precision), e.g., precision 2 → step 0.01
			return Math.pow(10, -attribute.Precision);
		}

		return undefined;
	};

	// Round value to the attribute's precision
	const roundToPrecision = (val: number): number => {
		const attrType = attribute.AttributeType;

		// Integer and BigInt: round to whole number
		if (attrType === "Integer" || attrType === "BigInt") {
			return Math.round(val);
		}

		// Decimal and Double: round to precision
		if ((attrType === "Decimal" || attrType === "Double") && attribute.Precision !== undefined) {
			const multiplier = Math.pow(10, attribute.Precision);
			return Math.round(val * multiplier) / multiplier;
		}

		return val;
	};

	// Handle input change
	const handleChange: React.ChangeEventHandler<HTMLInputElement> = (ev) => {
		const newValue = ev.target.value;
		setInputValue(newValue);

		// Parse and validate
		if (newValue === "") {
			onChange(undefined);
			return;
		}

		const parsed = parseFloat(newValue);
		if (!isNaN(parsed)) {
			// Round to precision to avoid floating-point errors
			const rounded = roundToPrecision(parsed);
			onChange(rounded);
		}
	};

	// Handle blur to validate min/max
	const handleBlur = () => {
		if (inputValue === "") {
			onChange(undefined);
			return;
		}

		const parsed = parseFloat(inputValue);
		if (isNaN(parsed)) {
			// Invalid number, clear
			setInputValue("");
			onChange(undefined);
			return;
		}

		// Round to precision first to avoid floating-point errors
		let constrainedValue = roundToPrecision(parsed);

		// Enforce min/max constraints
		if (attribute.MinValue !== undefined && constrainedValue < attribute.MinValue) {
			constrainedValue = attribute.MinValue;
		}

		if (attribute.MaxValue !== undefined && constrainedValue > attribute.MaxValue) {
			constrainedValue = attribute.MaxValue;
		}

		// Update if constrained or rounded
		if (constrainedValue !== parsed) {
			setInputValue(String(constrainedValue));
			onChange(constrainedValue);
		}
	};

	return (
		<Input
			id={inputId}
			type="number"
			value={inputValue}
			onChange={handleChange}
			onBlur={handleBlur}
			placeholder={placeholder}
			disabled={disabled}
			min={attribute.MinValue}
			max={attribute.MaxValue}
			step={getStep()}
		/>
	);
}
