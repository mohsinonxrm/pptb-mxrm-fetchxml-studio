/**
 * Bulleted list of parser findings, used inside MessageBars and the Parse to Tree dialog (#89).
 * A bar that appears is read as one piece of text, so each finding ends with a visually hidden
 * pause (a full stop, or a space after its own punctuation) and the findings don't run together
 * (#186 PR 3).
 */

import { makeStyles, tokens } from "@fluentui/react-components";

export interface ParseIssueItem {
	/** Element the finding refers to, e.g. "order" (rendered as <order>). */
	element?: string;
	message: string;
}

const useStyles = makeStyles({
	list: {
		marginTop: tokens.spacingVerticalXXS,
		marginBottom: 0,
		paddingLeft: tokens.spacingHorizontalL,
	},
	item: {
		marginTop: tokens.spacingVerticalXXS,
	},
	element: {
		fontFamily: tokens.fontFamilyMonospace,
		marginRight: tokens.spacingHorizontalXS,
	},
	hidden: {
		position: "absolute",
		width: "1px",
		height: "1px",
		margin: "-1px",
		padding: 0,
		overflow: "hidden",
		clip: "rect(0px, 0px, 0px, 0px)",
		whiteSpace: "nowrap",
		border: 0,
	},
});

export function ParseIssueList({ items }: { items: ParseIssueItem[] }) {
	const styles = useStyles();
	return (
		<ul className={styles.list}>
			{items.map((item, index) => (
				<li key={index} className={styles.item}>
					{item.element && <code className={styles.element}>&lt;{item.element}&gt;</code>}
					{item.message}
					<span className={styles.hidden}>
						{/[.!?…]$/.test(item.message.trim()) ? " " : ". "}
					</span>
				</li>
			))}
		</ul>
	);
}
