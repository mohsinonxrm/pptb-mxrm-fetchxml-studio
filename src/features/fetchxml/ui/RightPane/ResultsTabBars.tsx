/**
 * The Results tab's bars (#277), in one group under its command bar: the run that failed, the
 * export that failed or is running, and the paging notice. Each bar is about this tab, so it sits
 * here rather than above every tab (owner decision 6). Most severe first.
 */
import {
	Button,
	MessageBar,
	MessageBarActions,
	MessageBarBody,
	MessageBarGroup,
	MessageBarTitle,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { ArrowClockwise16Regular, Copy16Regular, Dismiss16Regular } from "@fluentui/react-icons";
import type { ReactElement } from "react";
import type { PagingNotice } from "../../api/paging";
import { useCopy } from "../../../../shared/feedback/useCopy";

const useStyles = makeStyles({
	container: {
		paddingLeft: "16px",
		paddingRight: "16px",
		paddingTop: "8px",
		paddingBottom: "8px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
	},
});

/** Why an export failed, and which export Try again repeats (#277) */
export interface ExportError {
	kind: "server" | "local";
	message: string;
	/** False when the export couldn't start (no privilege, no saved view, no rows): trying again won't help */
	retryable: boolean;
}

export interface ResultsTabBarsProps {
	/** Why the last run failed (Dataverse's message) */
	executeError?: string;
	/** Runs the query again, as Execute does */
	onRetryRun?: () => void;
	onDismissExecuteError?: () => void;
	isExporting?: boolean;
	exportError?: ExportError;
	onRetryExport?: () => void;
	onDismissExportError?: () => void;
	pagingNotice?: PagingNotice | null;
	/** Rows loaded so far, for the paging notice */
	rowCount: number;
	onRetryPaging?: () => void;
	onDismissPagingNotice?: () => void;
}

const DismissButton = ({ onClick }: { onClick?: () => void }) => (
	<Button
		appearance="transparent"
		icon={<Dismiss16Regular />}
		onClick={onClick}
		aria-label="Dismiss"
	/>
);

export function ResultsTabBars({
	executeError,
	onRetryRun,
	onDismissExecuteError,
	isExporting,
	exportError,
	onRetryExport,
	onDismissExportError,
	pagingNotice,
	rowCount,
	onRetryPaging,
	onDismissPagingNotice,
}: ResultsTabBarsProps) {
	const styles = useStyles();
	const copy = useCopy();

	const bars: ReactElement[] = [];

	if (executeError) {
		bars.push(
			<MessageBar key="run" intent="error" layout="multiline">
				<MessageBarBody>
					<MessageBarTitle>The query didn't run</MessageBarTitle>
					{executeError}
				</MessageBarBody>
				<MessageBarActions
					containerAction={<DismissButton onClick={onDismissExecuteError} />}
				>
					{onRetryRun && (
						<Button icon={<ArrowClockwise16Regular />} onClick={onRetryRun}>
							Try again
						</Button>
					)}
					<Button
						icon={<Copy16Regular />}
						onClick={() => void copy(executeError, "error details")}
					>
						Copy details
					</Button>
				</MessageBarActions>
			</MessageBar>,
		);
	}

	if (exportError && !isExporting) {
		bars.push(
			<MessageBar key="export-error" intent="error">
				<MessageBarBody>
					<MessageBarTitle>Export failed</MessageBarTitle>
					{exportError.message}
				</MessageBarBody>
				<MessageBarActions
					containerAction={<DismissButton onClick={onDismissExportError} />}
				>
					{exportError.retryable && onRetryExport && (
						<Button icon={<ArrowClockwise16Regular />} onClick={onRetryExport}>
							Try again
						</Button>
					)}
				</MessageBarActions>
			</MessageBar>,
		);
	}

	if (pagingNotice) {
		const kind = pagingNotice.kind;
		bars.push(
			<MessageBar
				key="paging"
				intent={
					kind === "error"
						? "error"
						: kind === "stopped" || kind === "limit"
							? "warning"
							: "info"
				}
				layout="multiline"
			>
				<MessageBarBody>
					<MessageBarTitle>
						{kind === "error"
							? "Couldn't load all records"
							: kind === "stopped"
								? "Retrieve all stopped"
								: kind === "limit"
									? `Retrieve all stopped at ${rowCount.toLocaleString()} rows`
									: "Paging without a paging cookie"}
					</MessageBarTitle>
					{pagingNotice.kind === "error"
						? `Loaded ${rowCount} rows; ${pagingNotice.page ? `page ${pagingNotice.page}` : "the next page"} failed: ${pagingNotice.message}`
						: pagingNotice.kind === "stopped"
							? `Stopped after ${rowCount} rows. Scroll to the end of the grid to load more, or execute again.`
							: pagingNotice.message}
				</MessageBarBody>
				<MessageBarActions
					containerAction={<DismissButton onClick={onDismissPagingNotice} />}
				>
					{kind === "error" && onRetryPaging && (
						<Button onClick={onRetryPaging}>Retry</Button>
					)}
				</MessageBarActions>
			</MessageBar>,
		);
	}

	if (isExporting) {
		bars.push(
			<MessageBar key="exporting" intent="info">
				<MessageBarBody>
					<MessageBarTitle>Exporting</MessageBarTitle>
					Exporting view to Excel…
				</MessageBarBody>
			</MessageBar>,
		);
	}

	if (bars.length === 0) return null;
	return (
		<div className={styles.container}>
			<MessageBarGroup>{bars}</MessageBarGroup>
		</div>
	);
}
