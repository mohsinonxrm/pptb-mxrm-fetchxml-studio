/**
 * A MessageBar's title, with a visually hidden full stop after it (#186 PR 3). A MessageBar
 * announces its text when it appears, and the space after Fluent's title is only CSS, so the title
 * and the text ran together: "The query didn't runAn unexpected error…". With the stop it reads
 * "The query didn't run. An unexpected error…". A title that ends in its own punctuation gets a
 * space only. Looks exactly like MessageBarTitle.
 */
import { MessageBarTitle, makeStyles, type MessageBarTitleProps } from "@fluentui/react-components";

const useStyles = makeStyles({
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

export function BarTitle({ children, ...props }: MessageBarTitleProps) {
	const styles = useStyles();
	const ownStop = typeof children === "string" && /[.!?…:]$/.test(children.trim());
	return (
		<MessageBarTitle {...props}>
			{children}
			<span className={styles.hidden}>{ownStop ? " " : ". "}</span>
		</MessageBarTitle>
	);
}
