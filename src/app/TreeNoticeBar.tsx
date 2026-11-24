/**
 * The MessageBar above the query tree (#277): a title, one or more sentences, and Dismiss. A long
 * body (the changes a view lost when it loaded) scrolls inside the bar instead of pushing the tree
 * out of the pane.
 */
import {
	Button,
	MessageBar,
	MessageBarActions,
	MessageBarBody,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { BarTitle } from "../shared/components/BarTitle";
import { DismissRegular } from "@fluentui/react-icons";
import type { TreeNoticeInput } from "./hooks/useTreeNotice";

const useStyles = makeStyles({
	bar: {
		flexShrink: 0,
		// Lined up with the Entity picker above it (12px either side)
		marginInline: tokens.spacingHorizontalM,
		marginBottom: tokens.spacingVerticalS,
	},
	body: {
		maxHeight: "160px",
		overflowY: "auto",
	},
});

export function TreeNoticeBar({
	notice,
	onDismiss,
}: {
	notice: TreeNoticeInput | null;
	onDismiss: () => void;
}) {
	const styles = useStyles();
	if (!notice) return null;
	return (
		<MessageBar intent={notice.intent} layout="multiline" className={styles.bar}>
			{/* Focusable, so a long list can be scrolled with the keyboard */}
			<MessageBarBody className={styles.body} tabIndex={0}>
				<BarTitle>{notice.title}</BarTitle>
				{notice.body}
			</MessageBarBody>
			<MessageBarActions
				containerAction={
					<Button
						appearance="transparent"
						aria-label="Dismiss"
						icon={<DismissRegular />}
						onClick={onDismiss}
					/>
				}
			/>
		</MessageBar>
	);
}
