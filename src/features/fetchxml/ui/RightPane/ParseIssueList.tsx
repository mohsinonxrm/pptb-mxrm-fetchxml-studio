/**
 * Bulleted list of parser findings, used inside MessageBars and the Parse to Tree dialog (#89).
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
});

export function ParseIssueList({ items }: { items: ParseIssueItem[] }) {
	const styles = useStyles();
	return (
		<ul className={styles.list}>
			{items.map((item, index) => (
				<li key={index} className={styles.item}>
					{item.element && <code className={styles.element}>&lt;{item.element}&gt;</code>}
					{item.message}
				</li>
			))}
		</ul>
	);
}
