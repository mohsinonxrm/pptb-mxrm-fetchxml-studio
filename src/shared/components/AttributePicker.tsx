/**
 * Shared Attribute Picker component using Fluent UI v9 Combobox
 * Loads attributes from metadata and displays as searchable dropdown
 */

import { useState, useEffect } from "react";
import { Combobox, useComboboxFilter, useId, type ComboboxProps } from "@fluentui/react-components";
import { loadEntityAttributes } from "../../features/fetchxml/api/dataverseMetadata";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";

interface AttributePickerProps {
	entityLogicalName: string;
	value: string;
	onChange: (logicalName: string) => void;
	placeholder?: string;
	disabled?: boolean;
}

export function AttributePicker({
	entityLogicalName,
	value,
	onChange,
	placeholder = "Select or type attribute name",
	disabled = false,
}: AttributePickerProps) {
	const comboId = useId("attribute-combobox");
	const [attributes, setAttributes] = useState<AttributeMetadata[]>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string | null>(null);
	const [query, setQuery] = useState<string>("");

	// Load attributes when entity changes
	useEffect(() => {
		if (!entityLogicalName) {
			setAttributes([]);
			return;
		}

		setLoading(true);
		setError(null);

		loadEntityAttributes(entityLogicalName)
			.then((attrs: AttributeMetadata[]) => {
				setAttributes(attrs);
				setLoading(false);
			})
			.catch((err: unknown) => {
				console.error("Failed to load attributes:", err);
				setError("Failed to load attributes");
				setLoading(false);
			});
	}, [entityLogicalName]);

	// Sync query with value when value changes or attributes load
	useEffect(() => {
		if (value && value !== "new_attribute" && attributes.length > 0) {
			const attr = attributes.find((a) => a.LogicalName === value);
			if (attr) {
				const displayName = attr.DisplayName?.UserLocalizedLabel?.Label || attr.SchemaName;
				setQuery(`${displayName} (${attr.LogicalName})`);
			} else {
				// If attribute not found in metadata, just show the logical name
				setQuery(value);
			}
		} else if (!value || value === "new_attribute") {
			setQuery("");
		}
	}, [value, attributes]);

	// Convert attributes to options for filtering
	const options = attributes.map((attr) => {
		const displayName = attr.DisplayName?.UserLocalizedLabel?.Label || attr.SchemaName;
		const label = `${displayName} (${attr.LogicalName})`;
		return {
			children: label,
			value: attr.LogicalName,
		};
	});

	// Use Fluent UI filtering
	const filteredChildren = useComboboxFilter(query, options, {
		noOptionsMessage: "No attributes match your search.",
	});

	const onOptionSelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		if (data.optionValue) {
			setQuery(data.optionText ?? "");
			onChange(data.optionValue);
		} else {
			// Clear button clicked - clear the attribute
			setQuery("");
			onChange("");
		}
	};

	// Track the selected option for visual indicator (but don't populate the input)
	const selectedOptions = value && value !== "new_attribute" ? [value] : [];

	if (error) {
		// Fallback to regular input on error
		return (
			<Combobox
				aria-labelledby={comboId}
				value={query}
				onChange={(ev) => {
					const inputValue = ev.target.value;
					setQuery(inputValue);
					onChange(inputValue);
				}}
				placeholder={placeholder}
				disabled={disabled}
			/>
		);
	}

	return (
		<Combobox
			aria-labelledby={comboId}
			onOptionSelect={onOptionSelect}
			selectedOptions={selectedOptions}
			value={query}
			onChange={(ev) => setQuery(ev.target.value)}
			placeholder={loading ? "Loading attributes..." : placeholder}
			disabled={disabled || loading}
			clearable
		>
			{filteredChildren}
		</Combobox>
	);
}
