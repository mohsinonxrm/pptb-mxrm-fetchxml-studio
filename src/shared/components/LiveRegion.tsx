/**
 * A visually hidden polite live region (#192): screen readers read `message` when the user is idle,
 * without moving focus. Render it once and keep it mounted, so the region exists before its text
 * changes (a region added together with its text isn't read).
 *
 * Fluent's AriaLiveAnnouncer isn't used here: mounted app-wide it makes every MessageBar announce
 * itself, and without `ariaNotify` its region is always assertive. #186 decides on that separately.
 */
import { makeStyles } from "@fluentui/react-components";

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

export function LiveRegion({ message }: { message: string }) {
	const styles = useStyles();
	return (
		<div role="status" aria-live="polite" aria-atomic="true" className={styles.hidden}>
			{message}
		</div>
	);
}
