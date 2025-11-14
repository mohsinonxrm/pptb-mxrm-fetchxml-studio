/**
 * Shared OptionSet Value Picker component using Fluent UI v9 Combobox
 * Loads option values from attribute metadata and displays as searchable dropdown
 * Supports Picklist, State, and Status attribute types
 */

import { useState, useEffect } from "react";
import { Combobox, useComboboxFilter, useId } from "@fluentui/react-components";
import { loadAttributeWithOptionSet } from "../../features/fetchxml/api/dataverseMetadata";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";

interface OptionSetOption {
	value: number;
	label: string;
	displayText: string; // "Label - Value" format
}

interface OptionSetValuePickerProps {
	entityLogicalName: string;
	attributeLogicalName: string;
	value?: number | string; // Can be number or string representation
	onChange: (value: number | undefined) => void;
	placeholder?: string;
	disabled?: boolean;
}

export function OptionSetValuePicker({
	entityLogicalName,
	attributeLogicalName,
	value,
	onChange,
	placeholder = "Select a value",
	disabled = false,
}: OptionSetValuePickerProps) {
	const comboId = useId("optionset-combobox");
	const [options, setOptions] = useState<OptionSetOption[]>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string | null>(null);
	const [query, setQuery] = useState<string>("");

	// Load option set values when entity/attribute changes
	useEffect(() => {
		if (!entityLogicalName || !attributeLogicalName) {
			setOptions([]);
			return;
		}

		setLoading(true);
		setError(null);

		loadAttributeWithOptionSet(entityLogicalName, attributeLogicalName)
			.then((attr: AttributeMetadata) => {
				if (!attr.OptionSet?.Options) {
					setError("No options found for this attribute");
					setOptions([]);
					setLoading(false);
					return;
				}

				// Convert metadata options to our format
				const optionsList: OptionSetOption[] = attr.OptionSet.Options.map((opt) => {
					const label = opt.Label?.UserLocalizedLabel?.Label || `Option ${opt.Value}`;
					return {
						value: opt.Value,
						label,
						displayText: `${label} - ${opt.Value}`,
					};
				});

				setOptions(optionsList);
				setLoading(false);
			})
			.catch((err: unknown) => {
				console.error("Failed to load option set values:", err);
				setError("Failed to load options");
				setLoading(false);
			});
	}, [entityLogicalName, attributeLogicalName]);

	// Sync query with value when value changes or options load
	useEffect(() => {
		if (value !== undefined && value !== null && value !== "") {
			const numValue = typeof value === "string" ? parseInt(value, 10) : value;
			const opt = options.find((o) => o.value === numValue);
			if (opt) {
				setQuery(opt.displayText);
			} else {
				setQuery(String(value));
			}
		} else {
			setQuery("");
		}
	}, [value, options]);

	// Convert options to combobox format and filter by both label and value
	const comboboxOptions = options
		.filter((opt) => {
			if (!query) return true;
			const searchLower = query.toLowerCase();
			return (
				opt.label.toLowerCase().includes(searchLower) || String(opt.value).includes(searchLower)
			);
		})
		.map((opt) => ({
			children: opt.displayText,
			value: String(opt.value), // Combobox uses string values
		}));

	// Apply Fluent UI's visual filtering only (we've already filtered the data)
	const filteredChildren = useComboboxFilter(query, comboboxOptions, {
		noOptionsMessage: loading ? "Loading options..." : error || "No options found",
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
