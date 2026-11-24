/**
 * Why a control can't be used, in a Tooltip that describes it (#279). The control stays focusable,
 * so keyboard and screen reader users find it and hear the reason: give a Button
 * `disabledFocusable`, a MenuItem `disabled` (already focusable in Fluent v9).
 *
 * The control is always wrapped the same way, with or without a reason: a focused Save that starts
 * saving gains a reason, and a different element tree would remount the button and drop focus to the
 * page. Without a reason the Tooltip never shows and describes nothing.
 *
 * A trigger around it (MenuTrigger) clones its child to add props and a ref; they're passed on to
 * the control.
 *
 * For an input, the reason goes in its Field's `hint` instead (the Fluent pattern).
 */
import type { ReactElement } from "react";
import type { TooltipProps } from "@fluentui/react-components";
import { StableTooltip } from "./StableTooltip";

export function withReason(
	element: ReactElement,
	/** Why the control can't be used; nothing when it can */
	reason: string | false | null | undefined,
	/** Where the Tooltip shows; menu items use "after" */
	positioning?: TooltipProps["positioning"],
): ReactElement {
	return (
		<StableTooltip content={reason} positioning={positioning}>
			{element}
		</StableTooltip>
	);
}
