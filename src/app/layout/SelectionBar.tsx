/**
 * The builder pane's selection bar (#234): while Properties live in the inspector drawer, it names
 * the selected node ("Condition · revenue") and opens its properties.
 */
import { forwardRef } from "react";
import { Button, Caption1, Tooltip, makeStyles, tokens } from "@fluentui/react-components";
import { Edit20Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
	bar: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}`,
		borderTop: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
		flexShrink: 0,
	},
	label: {
		flex: 1,
		minWidth: 0,
		overflow: "hidden",
		whiteSpace: "nowrap",
		textOverflow: "ellipsis",
		color: tokens.colorNeutralForeground2,
	},
});

export interface SelectionBarProps {
	/** "Kind · name", or "Nothing selected" */
	label: string;
	hasSelection: boolean;
	onEdit: () => void;
}

export const SelectionBar = forwardRef<HTMLButtonElement, SelectionBarProps>(function SelectionBar(
	{ label, hasSelection, onEdit },
	ref,
) {
	const styles = useStyles();
	const button = (
		<Button
			ref={ref}
			size="small"
			icon={<Edit20Regular />}
			disabledFocusable={!hasSelection}
			onClick={hasSelection ? onEdit : undefined}
		>
			Edit properties
		</Button>
	);
	return (
		<div className={styles.bar} role="region" aria-label="Selected node">
			<Caption1 className={styles.label} title={label}>
				{label}
			</Caption1>
			{hasSelection ? (
				button
			) : (
				<Tooltip content="Select a node in the tree first" relationship="description">
					{button}
				</Tooltip>
			)}
		</div>
	);
});
