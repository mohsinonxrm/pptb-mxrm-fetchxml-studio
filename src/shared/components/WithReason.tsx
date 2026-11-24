/**
 * Why a control can't be used, in a Tooltip that describes it (#279). The control stays focusable,
 * so keyboard and screen reader users find it and hear the reason: give a Button
 * `disabledFocusable`, a MenuItem `disabled` (already focusable in Fluent v9). Without a reason the
 * element is returned as it is.
 *
 * A function, not a component: a trigger around it (MenuTrigger, DialogTrigger) clones its child to
 * add props, and Tooltip passes them on to the control; a wrapper component would swallow them.
 *
 * For an input, the reason goes in its Field's `hint` instead (the Fluent pattern).
 */
import type { ReactElement } from "react";
import { Tooltip, type TooltipProps } from "@fluentui/react-components";

export function withReason(
	element: ReactElement,
	/** Why the control can't be used; nothing when it can */
	reason: string | false | null | undefined,
	/** Where the Tooltip shows; menu items use "after" */
	positioning?: TooltipProps["positioning"],
): ReactElement {
	if (!reason) return element;
	return (
		<Tooltip content={reason} relationship="description" positioning={positioning}>
			{element}
		</Tooltip>
	);
}
