/**
 * Shared Boolean Value Picker component using Fluent UI v9 Combobox
 * Loads true/false labels from TwoOptions attribute metadata
 * Displays custom labels like "Yes - true" / "No - false" or "On - true" / "Off - false"
 */

import { useState, useEffect } from "react";
import { Combobox, useComboboxFilter, useId } from "@fluentui/react-components";
import { loadAttributeWithOptionSet } from "../../features/fetchxml/api/dataverseMetadata";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";

interface BooleanOption {
	value: boolean;
	label: string;
	displayText: string; // "Label - value" format
}

interface BooleanValuePickerProps {
	entityLogicalName: string;
	attributeLogicalName: string;
	value?: boolean | string; // Can be boolean, "true"/"false", or "1"/"0"
	onChange: (value: boolean | undefined) => void;
	placeholder?: string;
	disabled?: boolean;
}

export function BooleanValuePicker({
	entityLogicalName,
	attributeLogicalName,
	value,
	onChange,
	placeholder = "Select true or false",
	disabled = false,
}: BooleanValuePickerProps) {
	const comboId = useId("boolean-combobox");
	const [options, setOptions] = useState<BooleanOption[]>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string | null>(null);
	const [query, setQuery] = useState<string>("");

	// Load boolean labels from attribute metadata
	useEffect(() => {
		if (!entityLogicalName || !attributeLogicalName) {
			setOptions([]);
			return;
		}

		setLoading(true);
		setError(null);

		loadAttributeWithOptionSet(entityLogicalName, attributeLogicalName)
			.then((attr: AttributeMetadata) => {
				if (!attr.OptionSet) {
					// Fallback to default labels if metadata doesn't have OptionSet
					console.warn("No OptionSet found for boolean attribute, using defaults");
					setOptions([
						{ value: true, label: "Yes", displayText: "Yes - true" },
						{ value: false, label: "No", displayText: "No - false" },
					]);
					setLoading(false);
					return;
				}

				// TwoOptions attributes have TrueOption and FalseOption
				const trueLabel = attr.OptionSet.TrueOption?.Label?.UserLocalizedLabel?.Label || "Yes";
				const falseLabel = attr.OptionSet.FalseOption?.Label?.UserLocalizedLabel?.Label || "No";

				setOptions([
					{ value: true, label: trueLabel, displayText: `${trueLabel} - true` },
					{ value: false, label: falseLabel, displayText: `${falseLabel} - false` },
				]);
				setLoading(false);
			})
			.catch((err: unknown) => {
				console.error("Failed to load boolean attribute metadata:", err);
				// Fallback to defaults on error
				setOptions([
					{ value: true, label: "Yes", displayText: "Yes - true" },
					{ value: false, label: "No", displayText: "No - false" },
				]);
				setLoading(false);
			});
	}, [entityLogicalName, attributeLogicalName]);

	// Sync query with value when value changes or options load
	useEffect(() => {
		if (value !== undefined && value !== null && value !== "") {
			// Normalize value to boolean
			let boolValue: boolean;
			if (typeof value === "boolean") {
				boolValue = value;
			} else if (typeof value === "string") {
				boolValue = value === "true" || value === "1";
			} else {
				boolValue = Boolean(value);
			}

			const opt = options.find((o) => o.value === boolValue);
			if (opt) {
				setQuery(opt.displayText);
			} else {
				setQuery(String(value));
			}
		} else {
			setQuery("");
		}
	}, [value, options]);

	// Convert options to combobox format and filter by both label and boolean value
	const comboboxOptions = options
		.filter((opt) => {
			if (!query) return true;
			const searchLower = query.toLowerCase();
			return (
				opt.label.toLowerCase().includes(searchLower) ||
				String(opt.value).toLowerCase().includes(searchLower)
			);
		})
		.map((opt) => ({
			children: opt.displayText,
			value: String(opt.value), // Combobox uses string values
		}));

	// Apply Fluent UI's visual filtering only (we've already filtered the data)
	const filteredChildren = useComboboxFilter(query, comboboxOptions, {
		noOptionsMessage: loading ? "Loading..." : error || "No options",
		filter: () => true, // Always return true since we pre-filtered
	});

	const handleChange: React.ChangeEventHandler<HTMLInputElement> = (ev) => {
		setQuery(ev.target.value);
	};

	const handleOptionSelect: NonNullable<React.ComponentProps<typeof Combobox>["onOptionSelect"]> = (
		_ev,
		data
	) => {
		if (!data.optionValue) {
			// Clear button clicked
			setQuery("");
			onChange(undefined);
			return;
		}

		const selectedOption = options.find((opt) => String(opt.value) === data.optionValue);
		if (selectedOption) {
			setQuery(selectedOption.displayText);
			onChange(selectedOption.value);
		}
	};

	return (
		<Combobox
			aria-labelledby={comboId}
			placeholder={placeholder}
			value={query}
			onChange={handleChange}
			onOptionSelect={handleOptionSelect}
			disabled={disabled || loading}
			clearable
		>
			{filteredChildren}
		</Combobox>
	);
}
