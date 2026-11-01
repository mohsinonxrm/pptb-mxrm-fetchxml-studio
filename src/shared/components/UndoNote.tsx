/**
 * The confirmation's way back (#69): the PPTB marketplace policy asks a confirmation to offer a way
 * back where feasible and, when there's none, to say so plainly.
 * https://docs.powerplatformtoolbox.com/policies/marketplace
 */
import type { ReactNode } from "react";
import { makeStyles, mergeClasses, Text, tokens } from "@fluentui/react-components";
import { ArrowUndo20Regular, Warning20Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
	row: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
	},
	none: {
		color: tokens.colorPaletteRedForeground1,
	},
	possible: {
		color: tokens.colorNeutralForeground2,
	},
	icon: {
		flexShrink: 0,
	},
});

export interface UndoNoteProps {
	/** "none": nothing can undo it (shown like Delete's warning); "possible": the text says how */
	undo: "none" | "possible";
	children: ReactNode;
}

export function UndoNote({ undo, children }: UndoNoteProps) {
	const styles = useStyles();
	const Icon = undo === "none" ? Warning20Regular : ArrowUndo20Regular;
	return (
		<div className={mergeClasses(styles.row, undo === "none" ? styles.none : styles.possible)}>
			<Icon className={styles.icon} />
			<Text weight={undo === "none" ? "semibold" : "regular"}>{children}</Text>
		</div>
	);
}
