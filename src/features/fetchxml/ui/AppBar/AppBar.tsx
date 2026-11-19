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
import { type ReactNode } from "react";
import { MenuItem, makeStyles, tokens } from "@fluentui/react-components";
import {
	Copy20Regular,
	Play20Filled,
	Play20Regular,
	Settings20Filled,
	Settings20Regular,
	bundleIcon,
} from "@fluentui/react-icons";
import { OverflowToolbar, type BarCommand } from "../../../../shared/components/commandBar";
import { useCopy } from "../../../../shared/feedback/useCopy";
import { useSaveViewCommand, type SaveViewCommandOptions } from "../Toolbar/useSaveViewCommand";
import {
	useSendToToolCommand,
	type SendToToolCommandOptions,
} from "../Toolbar/useSendToToolCommand";
import { useReturnFetchXmlCommand } from "../Toolbar/useReturnFetchXmlCommand";

const RunIcon = bundleIcon(Play20Filled, Play20Regular);
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
});

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
	/** Before the commands: the view's identity (#278 PR 2) */
	start?: ReactNode;
}

export function AppBar({
	run,
	saveView,
	sendToTool,
	isCallee,
	getCurrentXml,
	onOpenSettings,
	start,
}: AppBarProps) {
	const styles = useStyles();
	const copy = useCopy();
	const save = useSaveViewCommand(saveView);
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
			hint: "Run query (Ctrl+Enter)",
			onClick: run.onClick,
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
				// Settings (and Operations, #296) at the right, as in a model-driven app
				endAtEdge
				start={start}
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
		</>
	);
}
