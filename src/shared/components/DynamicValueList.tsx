/**
 * Dynamic list of input fields with add/remove buttons for multi-value operators
 * Each value gets its own input field to handle values with commas, spaces, etc.
 */

import { useState, useEffect } from "react";
import { Input, Button, makeStyles, tokens } from "@fluentui/react-components";
import { Add20Regular, Dismiss20Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
		width: "100%",
	},
	row: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
	},
	input: {
		flexGrow: 1,
	},
	addButton: {
		marginTop: tokens.spacingVerticalXS,
	},
});

interface DynamicValueListProps {
	values: string[];
	onChange: (values: string[]) => void;
	placeholder?: string;
	inputType?: "text" | "number";
}

/**
 * Component that displays a dynamic list of input fields
 * User can add/remove fields as needed
 */
export function DynamicValueList({
	values,
	onChange,
	placeholder = "Enter value",
	inputType = "text",
}: DynamicValueListProps) {
	const styles = useStyles();
	const [internalValues, setInternalValues] = useState<string[]>(values.length > 0 ? values : [""]);

	// Sync external changes
	useEffect(() => {
		if (values.length === 0 && internalValues.length === 1 && internalValues[0] === "") {
			// Don't update if external is empty and we just have one empty field
			return;
		}

		// Only update if values actually changed
		const currentNonEmpty = internalValues.filter((v) => v.trim() !== "");
		const valuesChanged =
			values.length !== currentNonEmpty.length || values.some((v, i) => v !== currentNonEmpty[i]);

		if (valuesChanged) {
			setInternalValues(values.length > 0 ? values : [""]);
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [values]); // internalValues intentionally excluded to prevent infinite loop

	const handleValueChange = (index: number, newValue: string) => {
		const updated = [...internalValues];
		updated[index] = newValue;
		setInternalValues(updated);

		// Filter out empty values before notifying parent
		const nonEmpty = updated.filter((v) => v.trim() !== "");
		onChange(nonEmpty);
	};

	const handleAddField = () => {
		const updated = [...internalValues, ""];
		setInternalValues(updated);
	};

	const handleRemoveField = (index: number) => {
		if (internalValues.length === 1) {
			// Don't remove the last field, just clear it
			const updated = [""];
			setInternalValues(updated);
			onChange([]);
		} else {
			const updated = internalValues.filter((_, i) => i !== index);
			setInternalValues(updated);
			const nonEmpty = updated.filter((v) => v.trim() !== "");
			onChange(nonEmpty);
		}
	};

	return (
		<div className={styles.container}>
			{internalValues.map((value, index) => (
				<div key={index} className={styles.row}>
					<Input
						className={styles.input}
						value={value}
						onChange={(_, data) => handleValueChange(index, data.value)}
						placeholder={`${placeholder} ${index + 1}`}
						type={inputType}
					/>
					{internalValues.length > 1 && (
						<Button
							appearance="subtle"
							icon={<Dismiss20Regular />}
							onClick={() => handleRemoveField(index)}
							aria-label={`Remove value ${index + 1}`}
						/>
					)}
				</div>
			))}
			<Button
				className={styles.addButton}
				appearance="secondary"
				icon={<Add20Regular />}
				onClick={handleAddField}
			>
				Add Value
			</Button>
		</div>
	);
}
