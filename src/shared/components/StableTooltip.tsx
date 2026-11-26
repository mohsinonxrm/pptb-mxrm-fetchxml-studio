/**
 * A Tooltip that is always there (#279, #392), so the control keeps its place in the element tree
 * whether or not there's something to say. A focused control that gains a reason (Save starts
 * saving, Delete loses its selection) would otherwise be remounted, and focus would drop to the page.
 * With no content it never shows and names or describes nothing.
 *
 * Marked as a Fluent trigger component, like Tooltip itself: a MenuTrigger or OverflowItem around it
 * applies its props and ref to the control inside, not to this wrapper.
 */
import { useEffect, useState, type FC, type ReactElement } from "react";
import { Tooltip, type TooltipProps } from "@fluentui/react-components";

export interface StableTooltipProps {
	/** What the Tooltip says; nothing for no Tooltip */
	content: string | false | null | undefined;
	/** How it relates to the control (default: it describes it) */
	relationship?: "label" | "description";
	positioning?: TooltipProps["positioning"];
	children: ReactElement;
}

/** Fluent's trigger marker (FluentTriggerComponent in react-utilities) */
type TriggerComponent = { isFluentTriggerComponent?: boolean };

export const StableTooltip: FC<StableTooltipProps> & TriggerComponent = ({
	content,
	relationship = "description",
	positioning,
	children,
}) => {
	const [visible, setVisible] = useState(false);
	// A hover while there was nothing to say isn't remembered: the control gaining a reason later
	// mustn't pop its Tooltip up with nothing pointing at it (#279 review)
	useEffect(() => {
		if (!content) setVisible(false);
	}, [content]);
	return (
		<Tooltip
			content={content || ""}
			relationship={content ? relationship : "inaccessible"}
			positioning={positioning}
			visible={!!content && visible}
			onVisibleChange={(_event, data) => setVisible(!!content && data.visible)}
		>
			{children}
		</Tooltip>
	);
};
StableTooltip.isFluentTriggerComponent = true;
