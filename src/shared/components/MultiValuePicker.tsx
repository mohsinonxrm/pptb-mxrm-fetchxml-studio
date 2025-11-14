/**
 * Input picker for operators that require MULTIPLE values (e.g., in, not-in, contain-values)
 * Uses Fluent UI v9 TagPicker for multi-select picklists and DynamicValueList for other types
 */

import { useState, useEffect } from "react";
import {
	TagPicker,
	TagPickerControl,
	TagPickerGroup,
	TagPickerInput,
	TagPickerList,
	TagPickerOption,
	Tag,
	makeStyles,
} from "@fluentui/react-components";
import type { TagPickerProps } from "@fluentui/react-components";
import { loadAttributeWithOptionSet } from "../../features/fetchxml/api/dataverseMetadata";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";
import { DynamicValueList } from "./DynamicValueList";

const useStyles = makeStyles({
	container: {
		width: "100%",
	},
});

interface MultiValuePickerProps {
	entityLogicalName: string;
	attributeLogicalName: string;
	attribute: AttributeMetadata;
	value?: string[] | number[]; // Array of values
	onChange: (value: string[] | number[] | undefined) => void;
	placeholder?: string;
}

/**
 * MultiValuePicker component for multi-value operators like 'in', 'not-in', 'contain-values'
 */
export function MultiValuePicker({
	entityLogicalName,
	attributeLogicalName,
	attribute,
	value = [],
	onChange,
	placeholder = "Enter value",
}: MultiValuePickerProps) {
	const styles = useStyles();
	const [options, setOptions] = useState<Array<{ label: string; value: number }>>([]);
	const [selectedValues, setSelectedValues] = useState<string[]>((value ?? []).map(String));

	const isPicklist =
		attribute.AttributeType === "Picklist" ||
		attribute.AttributeType === "State" ||
		attribute.AttributeType === "Status";

	// Load picklist options when needed
	useEffect(() => {
		if (isPicklist) {
			loadAttributeWithOptionSet(entityLogicalName, attributeLogicalName)
				.then((metadata: AttributeMetadata) => {
					if (metadata.OptionSet?.Options) {
						const opts = metadata.OptionSet.Options.map((opt) => ({
							label: opt.Label?.UserLocalizedLabel?.Label || String(opt.Value),
							value: opt.Value,
						}));
						setOptions(opts);
					}
				})
				.catch((error: unknown) => {
					console.error("Failed to load OptionSet metadata:", error);
				});
		}
	}, [entityLogicalName, attributeLogicalName, isPicklist]);

	// Sync external value changes to internal state (for picklist only)
	useEffect(() => {
		if (!isPicklist) return;

		const newSelectedValues = (value ?? []).map(String);
		// Only update if the values actually changed (avoid infinite loop)
		const hasChanged =
			newSelectedValues.length !== selectedValues.length ||
			newSelectedValues.some((v, i) => v !== selectedValues[i]);

		if (hasChanged) {
			setSelectedValues(newSelectedValues);
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [value, isPicklist]); // selectedValues intentionally excluded to prevent infinite loop

	const handleDynamicListChange = (values: string[]) => {
		// Determine if we need to convert to numbers
		const isNumeric =
			attribute.AttributeType === "Integer" ||
			attribute.AttributeType === "BigInt" ||
			attribute.AttributeType === "Decimal" ||
			attribute.AttributeType === "Double" ||
			attribute.AttributeType === "Money";

		if (isNumeric) {
			const numbers = values.map((v) => parseFloat(v)).filter((n) => !isNaN(n));
			onChange(numbers.length > 0 ? numbers : undefined);
		} else {
			onChange(values.length > 0 ? values : undefined);
		}
	};

	const handlePicklistChange: TagPickerProps["onOptionSelect"] = (_, data) => {
		const selected = data.selectedOptions.map((opt) => parseInt(opt, 10)).filter((n) => !isNaN(n));

		// Update local state immediately for controlled component behavior
		setSelectedValues(data.selectedOptions);

		// Notify parent of the change (parent will update value prop, but useEffect will detect no change)
		onChange(selected.length > 0 ? selected : undefined);
	};

	if (isPicklist) {
		const availableOptions = options.filter((opt) => !selectedValues.includes(String(opt.value)));

		return (
			<div className={styles.container}>
				<TagPicker onOptionSelect={handlePicklistChange} selectedOptions={selectedValues}>
					<TagPickerControl>
						<TagPickerGroup aria-label="Selected values">
							{selectedValues.map((val) => {
								const opt = options.find((o) => String(o.value) === val);
								return (
									<Tag key={val} shape="rounded" value={val}>
										{opt ? `${opt.label} - ${opt.value}` : val}
									</Tag>
								);
							})}
						</TagPickerGroup>
						<TagPickerInput aria-label={placeholder} placeholder={placeholder} />
					</TagPickerControl>
					<TagPickerList>
						{availableOptions.length > 0 ? (
							availableOptions.map((opt) => (
								<TagPickerOption
									key={opt.value}
									value={String(opt.value)}
									text={`${opt.label} - ${opt.value}`}
								>
									{opt.label} - {opt.value}
								</TagPickerOption>
							))
						) : (
							<TagPickerOption value="no-options" text="No options available">
								No options available
							</TagPickerOption>
						)}
					</TagPickerList>
				</TagPicker>
			</div>
		);
	}

	// For non-picklist attributes, use dynamic value list with add/remove buttons
	const isNumeric =
		attribute.AttributeType === "Integer" ||
		attribute.AttributeType === "BigInt" ||
		attribute.AttributeType === "Decimal" ||
		attribute.AttributeType === "Double" ||
		attribute.AttributeType === "Money";

	return (
		<div className={styles.container}>
			<DynamicValueList
				values={(value ?? []).map(String)}
				onChange={handleDynamicListChange}
				placeholder={placeholder}
				inputType={isNumeric ? "number" : "text"}
			/>
		</div>
	);
}
