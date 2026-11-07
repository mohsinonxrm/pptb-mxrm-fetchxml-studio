/**
 * Quick find test value (#142): shown while the query has {0} placeholders, as quick find views
 * store them. The value replaces {0} in the FetchXML that runs; the query itself keeps {0}, so
 * saving and copying are unchanged (savedViews.ts applyQuickFindValue).
 * https://learn.microsoft.com/power-apps/developer/data-platform/quick-find
 */

import { Checkbox, Field, Input, Text, makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
	panel: {
		display: "flex",
		flexDirection: "column",
		gap: "8px",
		padding: "10px 12px",
		backgroundColor: tokens.colorNeutralBackground3,
		borderRadius: tokens.borderRadiusMedium,
	},
	input: {
		maxWidth: "360px",
	},
	code: {
		fontFamily: tokens.fontFamilyMonospace,
	},
});

interface QuickFindTestValueProps {
	/** What the user typed */
	value: string;
	onValueChange: (value: string) => void;
	/** Add % at the end when it's missing: like "requires a search string that ends with %" */
	addPercent: boolean;
	onAddPercentChange: (addPercent: boolean) => void;
	/** The value that runs (toQuickFindSearchValue), shown once there is one */
	searchValue: string;
	disabled?: boolean;
}

export function QuickFindTestValue({
	value,
	onValueChange,
	addPercent,
	onAddPercentChange,
	searchValue,
	disabled,
}: QuickFindTestValueProps) {
	const styles = useStyles();
	const placeholder = <span className={styles.code}>{"{0}"}</span>;
	const missing = !value.trim();
	return (
		<div className={styles.panel}>
			<Field
				label="Quick find value"
				hint={
					<>
						Replaces {placeholder} when you execute. The query keeps {placeholder}, so saving and copying are
						unchanged.
					</>
				}
				validationState={missing ? "warning" : "none"}
				validationMessage={missing ? "Enter a value to execute this quick find view." : undefined}
			>
				<Input
					className={styles.input}
					value={value}
					onChange={(_, data) => onValueChange(data.value)}
					placeholder="e.g. contoso"
					disabled={disabled}
				/>
			</Field>
			<Checkbox
				label="Add % at the end"
				checked={addPercent}
				disabled={disabled}
				onChange={(_, data) => onAddPercentChange(data.checked === true)}
			/>
			{searchValue && (
				<Text size={200}>
					Runs with <span className={styles.code}>value="{searchValue}"</span> in each {placeholder} condition.
				</Text>
			)}
		</div>
	);
}
