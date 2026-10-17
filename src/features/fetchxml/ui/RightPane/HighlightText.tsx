/**
 * Text with the quick search's matches highlighted (#23)
 */

import { useContext, type ReactNode } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { SearchHighlightContext, searchMatches } from "./searchHighlightContext";

const useStyles = makeStyles({
	mark: {
		backgroundColor: tokens.colorPaletteYellowBackground2,
		color: "inherit",
		borderRadius: tokens.borderRadiusSmall,
	},
});

export function HighlightText({ text }: { text: string }) {
	const styles = useStyles();
	const search = useContext(SearchHighlightContext);
	const ranges = searchMatches(text, search);
	if (ranges.length === 0) return <>{text}</>;
	const parts: ReactNode[] = [];
	let at = 0;
	for (const [start, end] of ranges) {
		if (start > at) parts.push(text.slice(at, start));
		parts.push(
			<mark key={start} className={styles.mark}>
				{text.slice(start, end)}
			</mark>,
		);
		at = end;
	}
	if (at < text.length) parts.push(text.slice(at));
	return <>{parts}</>;
}
