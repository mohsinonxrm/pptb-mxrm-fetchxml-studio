/**
 * Confirmation when "Aggregate query" is turned off and aggregate-only settings would stop working
 * (#90b, D4). The plan comes from planAggregateOff (model/aggregateToggle.ts).
 */
import {
	Body1,
	Body1Strong,
	Button,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import type { AggregateOffPlan } from "../../../../model/aggregateToggle";
import { ParseIssueList } from "../../../RightPane/ParseIssueList";

const useStyles = makeStyles({
	content: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
});

interface AggregateOffDialogProps {
	plan: AggregateOffPlan | null;
	/** Turn off and apply the plan */
	onConfirm: () => void;
	/** Turn off but keep the settings (they then show as errors) */
	onTurnOffOnly: () => void;
	onCancel: () => void;
}

export function AggregateOffDialog({ plan, onConfirm, onTurnOffOnly, onCancel }: AggregateOffDialogProps) {
	const styles = useStyles();
	return (
		<Dialog modalType="alert" open={plan !== null} onOpenChange={(_, data) => !data.open && onCancel()}>
			<DialogSurface>
				<DialogBody>
					<DialogTitle>Turn off aggregate query?</DialogTitle>
					<DialogContent className={styles.content}>
						<Body1>
							These settings only work in aggregate queries, and Dataverse rejects some of them in a normal query.
						</Body1>
						{plan && plan.removed.length > 0 && (
							<div>
								<Body1Strong>Removed</Body1Strong>
								<ParseIssueList items={plan.removed} />
							</div>
						)}
						{plan && plan.converted.length > 0 && (
							<div>
								<Body1Strong>Changed</Body1Strong>
								<ParseIssueList items={plan.converted} />
							</div>
						)}
					</DialogContent>
					<DialogActions position="start">
						<Button appearance="secondary" onClick={onTurnOffOnly}>
							Turn off only
						</Button>
					</DialogActions>
					<DialogActions position="end">
						<Button appearance="secondary" onClick={onCancel}>
							Cancel
						</Button>
						<Button appearance="primary" onClick={onConfirm}>
							Turn off and update
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
