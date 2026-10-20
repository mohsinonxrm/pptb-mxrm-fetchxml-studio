/**
 * A row of toolbar actions that moves the actions which don't fit into a "More actions" menu.
 * Built on Fluent's Overflow: the row measures itself, and every action that is hidden shows up
 * as a menu item instead. Use it for the buttons next to a fixed control (e.g. a switch or Execute)
 * that must always stay visible.
 */

import type { ReactElement } from "react";
import {
	Button,
	Menu,
	MenuItem,
	MenuList,
	MenuPopover,
	MenuTrigger,
	Overflow,
	OverflowItem,
	Tooltip,
	ToolbarButton,
	makeStyles,
	tokens,
	useIsOverflowItemVisible,
	useOverflowMenu,
} from "@fluentui/react-components";
import { MoreHorizontal20Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
	row: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		minWidth: 0,
		overflow: "hidden",
	},
});

export interface OverflowAction {
	id: string;
	label: string;
	icon: ReactElement;
	onClick?: () => void;
	disabled?: boolean;
	/** Tooltip text; defaults to the label */
	tooltip?: string;
}

/** Menu entry for one action; only rendered while the action is hidden in the row */
function OverflowMenuEntry({ id, action }: { id: string; action: OverflowAction }) {
	const visible = useIsOverflowItemVisible(id);
	if (visible) return null;
	return (
		<MenuItem icon={action.icon} disabled={action.disabled} onClick={action.onClick}>
			{action.label}
		</MenuItem>
	);
}

function OverflowMenu({ actions }: { actions: OverflowAction[] }) {
	const { ref, isOverflowing } = useOverflowMenu<HTMLButtonElement>();
	if (!isOverflowing) return null;
	return (
		<Menu positioning="below-end">
			<MenuTrigger disableButtonEnhancement>
				<Button
					ref={ref}
					appearance="subtle"
					size="small"
					icon={<MoreHorizontal20Regular />}
					aria-label="More actions"
				/>
			</MenuTrigger>
			<MenuPopover>
				<MenuList>
					{actions.map((action) => (
						<OverflowMenuEntry key={action.id} id={action.id} action={action} />
					))}
				</MenuList>
			</MenuPopover>
		</Menu>
	);
}

export function OverflowActionBar({
	actions,
	label,
}: {
	actions: OverflowAction[];
	label: string;
}) {
	const styles = useStyles();
	return (
		<Overflow>
			<div className={styles.row} role="toolbar" aria-label={label}>
				{actions.map((action) => (
					<OverflowItem key={action.id} id={action.id}>
						<Tooltip content={action.tooltip ?? action.label} relationship="label">
							<ToolbarButton
								icon={action.icon}
								disabled={action.disabled}
								onClick={action.onClick}
							>
								{action.label}
							</ToolbarButton>
						</Tooltip>
					</OverflowItem>
				))}
				<OverflowMenu actions={actions} />
			</div>
		</Overflow>
	);
}
