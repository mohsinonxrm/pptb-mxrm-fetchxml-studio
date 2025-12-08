/**
 * SolutionPicker - Dropdown for selecting an unmanaged solution
 * Used when saving system views to allow adding them to a specific solution
 */

import { useMemo } from "react";
import { Combobox, Option, Field, makeStyles, tokens, Text } from "@fluentui/react-components";
import type { Solution } from "../../api/pptbClient";

const useStyles = makeStyles({
	container: {
		display: "flex",
		flexDirection: "column",
		gap: "4px",
	},
	combobox: {
		width: "100%",
	},
	hint: {
		fontSize: "12px",
		color: tokens.colorNeutralForeground3,
	},
	solutionOption: {
		display: "flex",
		flexDirection: "column",
		gap: "2px",
	},
	solutionName: {
		fontWeight: 500,
	},
	solutionUniqueName: {
		fontSize: "11px",
		color: tokens.colorNeutralForeground3,
	},
});

interface SolutionPickerProps {
	/** Available unmanaged solutions */
	solutions: Solution[];
	/** Currently selected solution */
	selectedSolution: Solution | null;
	/** Selection change callback */
	onSolutionChange: (solution: Solution | null) => void;
	/** Disabled state */
	disabled?: boolean;
}

export function SolutionPicker({
	solutions,
	selectedSolution,
	onSolutionChange,
	disabled = false,
}: SolutionPickerProps) {
	const styles = useStyles();

	// Sort solutions by friendly name
	const sortedSolutions = useMemo(() => {
		return [...solutions].sort((a, b) =>
			(a.friendlyname || a.uniquename).localeCompare(b.friendlyname || b.uniquename)
		);
	}, [solutions]);

	// Handle selection change
	const handleOptionSelect = (_: unknown, data: { optionValue?: string; optionText?: string }) => {
		if (data.optionValue === "__none__") {
			onSolutionChange(null);
		} else {
			const solution = solutions.find((s) => s.solutionid === data.optionValue);
			onSolutionChange(solution || null);
		}
	};

	return (
		<div className={styles.container}>
			<Field label="Add to Solution" hint="Select a solution to include the view in">
				<Combobox
					className={styles.combobox}
					placeholder="Select solution (optional)"
					value={selectedSolution?.friendlyname || ""}
					onOptionSelect={handleOptionSelect}
					disabled={disabled}
				>
					<Option value="__none__" text="">
						<span style={{ fontStyle: "italic", color: tokens.colorNeutralForeground3 }}>
							None - don't add to any solution
						</span>
					</Option>
					{sortedSolutions.map((solution) => (
						<Option
							key={solution.solutionid}
							value={solution.solutionid}
							text={solution.friendlyname || solution.uniquename}
						>
							<div className={styles.solutionOption}>
								<span className={styles.solutionName}>
									{solution.friendlyname || solution.uniquename}
								</span>
								<span className={styles.solutionUniqueName}>
									{solution.uniquename}
									{solution.version && ` (v${solution.version})`}
								</span>
							</div>
						</Option>
					))}
				</Combobox>
			</Field>
			{solutions.length === 0 && (
				<Text className={styles.hint}>
					No unmanaged solutions available. The view will be created in the Default solution.
				</Text>
			)}
		</div>
	);
}
