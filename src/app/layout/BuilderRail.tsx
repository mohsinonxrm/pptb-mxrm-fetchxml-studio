/**
 * The collapsed builder pane (#234): a 44 px rail with Expand, then Query tree and Properties, which
 * open that part in a peek over the output. Also the pane's own commands, beside the scope chip:
 * Collapse while docked, Pin while peeking.
 */
import { forwardRef } from "react";
import {
	Button,
	Divider,
	ToggleButton,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import {
	Options20Filled,
	Options20Regular,
	PanelLeftContract20Filled,
	PanelLeftContract20Regular,
	PanelLeftExpand20Filled,
	PanelLeftExpand20Regular,
	Pin20Filled,
	Pin20Regular,
	TextBulletListTree20Filled,
	TextBulletListTree20Regular,
	bundleIcon,
} from "@fluentui/react-icons";
import type { PeekPart } from "../hooks/useBuilderLayout";

const CollapseIcon = bundleIcon(PanelLeftContract20Filled, PanelLeftContract20Regular);
const ExpandIcon = bundleIcon(PanelLeftExpand20Filled, PanelLeftExpand20Regular);
const PinIcon = bundleIcon(Pin20Filled, Pin20Regular);
const TreeIcon = bundleIcon(TextBulletListTree20Filled, TextBulletListTree20Regular);
const PropertiesIcon = bundleIcon(Options20Filled, Options20Regular);

const useStyles = makeStyles({
	rail: {
		width: "44px",
		flexShrink: 0,
		display: "flex",
		flexDirection: "column",
		alignItems: "center",
		gap: tokens.spacingVerticalXS,
		paddingBlock: tokens.spacingVerticalS,
		borderRight: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
		backgroundColor: tokens.colorNeutralBackground2,
	},
	divider: { width: "24px", flexGrow: 0 },
});

export interface BuilderRailProps {
	peekPart: PeekPart | null;
	/** The peek's id, for aria-controls while it's open */
	peekId: string;
	onExpand: () => void;
	onTogglePeek: (part: PeekPart) => void;
	expandRef: React.Ref<HTMLButtonElement>;
	treeRef: React.Ref<HTMLButtonElement>;
	propertiesRef: React.Ref<HTMLButtonElement>;
	/** The rail itself, for Ctrl+F6 (#186) */
	navRef?: React.Ref<HTMLElement>;
}

export function BuilderRail({
	peekPart,
	peekId,
	onExpand,
	onTogglePeek,
	expandRef,
	treeRef,
	propertiesRef,
	navRef,
}: BuilderRailProps) {
	const styles = useStyles();
	return (
		<nav ref={navRef} className={styles.rail} aria-label="Builder pane">
			<Tooltip content="Expand the builder pane" relationship="label" positioning="after">
				<Button
					ref={expandRef}
					appearance="subtle"
					icon={<ExpandIcon />}
					onClick={onExpand}
				/>
			</Tooltip>
			<Divider className={styles.divider} />
			<Tooltip content="Query tree" relationship="label" positioning="after">
				<ToggleButton
					ref={treeRef}
					appearance="subtle"
					icon={<TreeIcon />}
					checked={peekPart === "tree"}
					aria-controls={peekPart === "tree" ? peekId : undefined}
					onClick={() => onTogglePeek("tree")}
				/>
			</Tooltip>
			<Tooltip content="Properties" relationship="label" positioning="after">
				<ToggleButton
					ref={propertiesRef}
					appearance="subtle"
					icon={<PropertiesIcon />}
					checked={peekPart === "props"}
					aria-controls={peekPart === "props" ? peekId : undefined}
					onClick={() => onTogglePeek("props")}
				/>
			</Tooltip>
		</nav>
	);
}

/** Collapse, beside the scope chip while the pane is docked */
export const CollapseButton = forwardRef<HTMLButtonElement, { onClick: () => void }>(
	function CollapseButton({ onClick }, ref) {
		return (
			<Tooltip content="Collapse the builder pane" relationship="label">
				<Button
					ref={ref}
					size="small"
					appearance="subtle"
					icon={<CollapseIcon />}
					onClick={onClick}
				/>
			</Tooltip>
		);
	},
);

/** Pin, in a peek: docks the pane again */
export const PinButton = forwardRef<HTMLButtonElement, { onClick: () => void }>(function PinButton(
	{ onClick },
	ref,
) {
	return (
		<Tooltip content="Pin the builder pane open" relationship="label">
			<Button
				ref={ref}
				size="small"
				appearance="subtle"
				icon={<PinIcon />}
				onClick={onClick}
			/>
		</Tooltip>
	);
});
