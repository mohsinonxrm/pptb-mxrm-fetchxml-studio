/**
 * The app bar (#278): one full-width row above both panes with the query's commands, in the
 * order of a model-driven app's command bar: Run | Save · Send to · Return FetchXML … More |
 * Settings. Run comes first, is the one primary button and never overflows; Settings never does
 * either. What doesn't fit moves into More by priority (D: Save 60, Send to 20, Return FetchXML
 * 10). More always has Copy FetchXML.
 *
 * Save and Send to keep their editor sync (EditorSyncContext), so AppShell renders this inside the
 * provider. New and Open view join in #278's next PRs; Operations in #296.
 */
import { useRef, useState } from "react";
import { Badge, Body1Strong, MenuItem, makeStyles, tokens } from "@fluentui/react-components";
import {
	Add20Filled,
	Add20Regular,
	Copy20Regular,
	Play20Filled,
	Play20Regular,
	Settings20Filled,
	Settings20Regular,
	bundleIcon,
} from "@fluentui/react-icons";
import { OverflowToolbar, type BarCommand } from "../../../../shared/components/commandBar";
import { formatKeys, getEntry } from "../../../../app/shortcuts/registry";
import { useCopy } from "../../../../shared/feedback/useCopy";
import { useSaveViewCommand, type SaveViewCommandOptions } from "../Toolbar/useSaveViewCommand";
import {
	useSendToToolCommand,
	type SendToToolCommandOptions,
} from "../Toolbar/useSendToToolCommand";
import { useReturnFetchXmlCommand } from "../Toolbar/useReturnFetchXmlCommand";
import { DiscardChangesDialog } from "../Dialogs/DiscardChangesDialog";
import { useNotify } from "../../../../shared/feedback/useNotify";

const RunIcon = bundleIcon(Play20Filled, Play20Regular);
const NewIcon = bundleIcon(Add20Filled, Add20Regular);
const SettingsIcon = bundleIcon(Settings20Filled, Settings20Regular);

/** Every group the bar can show, in bar order (OverflowToolbar learns them up front) */
const BAR_GROUPS = ["run", "file", "share", "settings"];

const useStyles = makeStyles({
	bar: {
		flexShrink: 0,
		paddingInline: tokens.spacingHorizontalS,
		paddingBlock: tokens.spacingVerticalXS,
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
		backgroundColor: tokens.colorNeutralBackground1,
	},
	identity: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		minWidth: 0,
		maxWidth: "320px",
		paddingInlineEnd: tokens.spacingHorizontalS,
	},
	viewName: {
		overflow: "hidden",
		whiteSpace: "nowrap",
		textOverflow: "ellipsis",
	},
	badge: { flexShrink: 0 },
});

/** What New says when it discards the query's changes */
export const discardedToast = (viewName: string | null) =>
	viewName ? `Discarded changes to "${viewName}"` : "Discarded the unsaved query";

export interface AppBarProps {
	run: {
		onClick: () => void;
		/** The query is running or more rows are loading */
		busy: boolean;
		disabledReason?: string;
	};
	saveView: SaveViewCommandOptions;
	/** Left out when no tool takes FetchXML */
	sendToTool?: SendToToolCommandOptions;
	/** Launched by another tool: Return FetchXML shows */
	isCallee?: boolean;
	/** The FetchXML the user sees (the editor's in editor mode) */
	getCurrentXml: () => string;
	onOpenSettings?: () => void;
	/** The view the query is (#278); null for "New query" */
	viewName: string | null;
	/** The query differs from the view as loaded or saved (its FetchXML, columns or widths) */
	unsaved: boolean;
	/** Start a new query: the caller resets the tree, the editor and the results */
	onNewQuery: () => void;
}

export function AppBar({
	run,
	saveView,
	sendToTool,
	isCallee,
	getCurrentXml,
	onOpenSettings,
	viewName,
	unsaved,
	onNewQuery,
}: AppBarProps) {
	const styles = useStyles();
	const copy = useCopy();
	const { notify } = useNotify();
	// New with unsaved changes asks first (#278); "Save and start new" starts new once the save is done
	const [confirmNew, setConfirmNew] = useState(false);
	const newAfterSave = useRef(false);
	const save = useSaveViewCommand({
		...saveView,
		onSaveComplete: (...args) => {
			saveView.onSaveComplete?.(...args);
			if (newAfterSave.current) {
				newAfterSave.current = false;
				onNewQuery();
			}
		},
		onCancel: () => {
			newAfterSave.current = false;
			saveView.onCancel?.();
		},
	});
	const handleNew = () => {
		if (unsaved) setConfirmNew(true);
		else onNewQuery();
	};
	const send = useSendToToolCommand(
		sendToTool ?? { fetchXml: "", entityLogicalName: "", tools: [] },
	);
	const returnCommand = useReturnFetchXmlCommand(getCurrentXml);

	const commands: BarCommand[] = [
		{
			id: "run",
			label: "Run",
			icon: RunIcon,
			group: "run",
			pinned: true,
			kind: "button",
			appearance: "primary",
			busy: run.busy,
			disabledReason: run.disabledReason,
			hint: `Run query (${formatKeys(getEntry("run")).join(" or ")})`,
			keyShortcuts: getEntry("run").aria,
			onClick: run.onClick,
		},
		{
			id: "new",
			label: "New",
			icon: NewIcon,
			group: "file",
			priority: 40,
			kind: "button",
			hint: "Start a new query",
			onClick: handleNew,
		},
		{ ...save.command, group: "file" },
	];
	if (send) commands.push({ ...send, group: "share" });
	if (isCallee) commands.push({ ...returnCommand, group: "share" });
	commands.push({
		id: "settings",
		label: "Settings",
		icon: SettingsIcon,
		group: "settings",
		placement: "end",
		kind: "button",
		iconOnly: true,
		onClick: onOpenSettings,
	});

	return (
		<>
			<OverflowToolbar
				ariaLabel="Query commands"
				commands={commands}
				groups={BAR_GROUPS}
				start={
					<div className={styles.identity}>
						<Body1Strong className={styles.viewName} title={viewName ?? "New query"}>
							{viewName ?? "New query"}
						</Body1Strong>
						{unsaved && (
							<Badge
								className={styles.badge}
								appearance="tint"
								color="warning"
								shape="rounded"
							>
								Unsaved changes
							</Badge>
						)}
					</div>
				}
				// Settings (and Operations, #296) at the right, as in a model-driven app
				endAtEdge
				className={styles.bar}
				moreExtra={
					<MenuItem
						icon={<Copy20Regular />}
						disabled={!getCurrentXml()}
						onClick={() => void copy(getCurrentXml(), "FetchXML")}
					>
						Copy FetchXML
					</MenuItem>
				}
			/>
			{save.dialog}
			<DiscardChangesDialog
				open={confirmNew}
				intent="new"
				viewName={viewName}
				onSave={() => {
					setConfirmNew(false);
					newAfterSave.current = true;
					save.command.onClick?.();
				}}
				onDiscard={() => {
					setConfirmNew(false);
					onNewQuery();
					notify.info(discardedToast(viewName));
				}}
				onCancel={() => setConfirmNew(false)}
			/>
		</>
	);
}
