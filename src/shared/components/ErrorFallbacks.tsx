/**
 * What the error boundaries show (#185): variant A of the reviewed prototype
 * (.claude/mockups/error-boundary), an error MessageBar with the actions.
 */
import {
	Accordion,
	AccordionHeader,
	AccordionItem,
	AccordionPanel,
	Button,
	Caption1,
	MessageBar,
	MessageBarActions,
	MessageBarBody,
	Text,
	Textarea,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { BarTitle } from "./BarTitle";
import {
	Add20Regular,
	ArrowClockwise20Regular,
	ArrowCounterclockwise20Regular,
	Copy20Regular,
} from "@fluentui/react-icons";
import { useCopy } from "../feedback/useCopy";

const useStyles = makeStyles({
	pane: {
		padding: tokens.spacingHorizontalS,
	},
	muted: {
		color: tokens.colorNeutralForeground3,
	},
	// Fills the tool's tab: nothing else is rendered when the whole tool fails
	pageWrap: {
		height: "100vh",
		overflow: "auto",
		display: "flex",
		justifyContent: "center",
		backgroundColor: tokens.colorNeutralBackground1,
	},
	page: {
		maxWidth: "720px",
		width: "100%",
		padding: tokens.spacingHorizontalXXL,
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalL,
	},
	actions: {
		display: "flex",
		gap: tokens.spacingHorizontalS,
		flexWrap: "wrap",
	},
	monospace: {
		fontFamily: tokens.fontFamilyMonospace,
		fontSize: tokens.fontSizeBase200,
	},
	details: {
		fontFamily: tokens.fontFamilyMonospace,
		fontSize: tokens.fontSizeBase200,
		whiteSpace: "pre-wrap",
		color: tokens.colorNeutralForeground2,
	},
});

/** Copies the text and says so in a toast: "Copied {what}" (#277) */
export function CopyButton({
	text,
	label,
	what,
	appearance,
}: {
	text: string;
	label: string;
	/** What the toast names: "Copied error details" */
	what: string;
	appearance?: "secondary" | "subtle";
}) {
	const copy = useCopy();
	return (
		<Button
			appearance={appearance}
			icon={<Copy20Regular />}
			onClick={() => void copy(text, what)}
		>
			{label}
		</Button>
	);
}

export function PaneErrorFallback({
	title,
	message,
	details,
	onTryAgain,
}: {
	title: string;
	message: string;
	details: string;
	onTryAgain: () => void;
}) {
	const styles = useStyles();
	return (
		<div className={styles.pane}>
			<MessageBar intent="error" layout="multiline">
				<MessageBarBody>
					<BarTitle>{title}</BarTitle>
					Something went wrong while showing this panel. The rest of FetchXML Studio still
					works. <Caption1 className={styles.muted}>({message})</Caption1>
				</MessageBarBody>
				<MessageBarActions>
					<Button icon={<ArrowClockwise20Regular />} onClick={onTryAgain}>
						Try again
					</Button>
					<CopyButton text={details} label="Copy details" what="error details" />
				</MessageBarActions>
			</MessageBar>
		</div>
	);
}

export function ToolErrorFallback({
	fetchXml,
	details,
	onRestart,
}: {
	fetchXml?: string;
	details: string;
	onRestart: (withQuery: boolean) => void;
}) {
	const styles = useStyles();
	return (
		<div className={styles.pageWrap}>
			<div className={styles.page}>
				<MessageBar intent="error" layout="multiline">
					<MessageBarBody>
						<BarTitle>FetchXML Studio hit a problem</BarTitle>
						{fetchXml
							? "Your query is kept. Restart to continue where you left off, or copy the FetchXML first."
							: "Start again to keep working."}
					</MessageBarBody>
				</MessageBar>
				{fetchXml && (
					<div>
						<Text weight="semibold" block>
							Your query
						</Text>
						<Textarea
							readOnly
							value={fetchXml}
							resize="vertical"
							textarea={{ className: styles.monospace, rows: 10 }}
							style={{ width: "100%" }}
							aria-label="Your query's FetchXML"
						/>
					</div>
				)}
				<div className={styles.actions}>
					{fetchXml && (
						<Button
							appearance="primary"
							icon={<ArrowCounterclockwise20Regular />}
							onClick={() => onRestart(true)}
						>
							Restart with this query
						</Button>
					)}
					<Button
						appearance={fetchXml ? "secondary" : "primary"}
						icon={<Add20Regular />}
						onClick={() => onRestart(false)}
					>
						Start a new query
					</Button>
					{fetchXml && (
						<CopyButton text={fetchXml} label="Copy FetchXML" what="FetchXML" />
					)}
					<CopyButton text={details} label="Copy details" what="error details" />
				</div>
				<Accordion collapsible>
					<AccordionItem value="details">
						<AccordionHeader size="small">Details</AccordionHeader>
						<AccordionPanel>
							<div className={styles.details}>{details}</div>
						</AccordionPanel>
					</AccordionItem>
				</Accordion>
			</div>
		</div>
	);
}
