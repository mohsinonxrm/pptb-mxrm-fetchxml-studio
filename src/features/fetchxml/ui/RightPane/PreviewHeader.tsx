/**
 * The right pane's header row (#232): the tabs, then Execute, Return FetchXML, Save and Send to
 * tool, then Settings. Both halves are shared overflow bars, so at any width nothing wraps or
 * clips: tabs move into "More", commands into "More commands". Execute and Settings never do.
 *
 * #278 moves the commands into the app bar; until then the two bars share this row.
 */
import { CounterBadge, makeStyles, tokens } from "@fluentui/react-components";
import {
	Code20Filled,
	Code20Regular,
	Play20Filled,
	Play20Regular,
	Settings20Filled,
	Settings20Regular,
	bundleIcon,
} from "@fluentui/react-icons";
import {
	OverflowTabList,
	OverflowToolbar,
	type BarCommand,
	type BarTab,
} from "../../../../shared/components/commandBar";
import type { PreviewTab } from "./previewHeaderRules";
import { useSaveViewCommand, type SaveViewCommandOptions } from "../Toolbar/useSaveViewCommand";
import {
	useSendToToolCommand,
	type SendToToolCommandOptions,
} from "../Toolbar/useSendToToolCommand";
import { useReturnFetchXmlCommand } from "../Toolbar/useReturnFetchXmlCommand";

const PlayIcon = bundleIcon(Play20Filled, Play20Regular);
const SettingsIcon = bundleIcon(Settings20Filled, Settings20Regular);
const CodeIcon = bundleIcon(Code20Filled, Code20Regular);

export type { PreviewTab } from "./previewHeaderRules";

/**
 * About how wide each part of the row is with nothing in More, in px (measured in Edge at 100%). The
 * two bars share the row in proportion to these, so the split doesn't depend on what has overflowed
 * (which would make them fight over the space).
 */
const TABS_WIDTH = 432;
const COMMAND_WIDTH: Record<string, number> = {
	execute: 104,
	return: 156,
	save: 128,
	send: 140,
	settings: 32,
};
/** A ToolbarDivider with its margins */
const DIVIDER_WIDTH = 24;
/** The space Overflow keeps free (COMMANDS_OVERFLOW_PADDING), plus the small toolbar's own padding */
const COMMANDS_OVERFLOW_PADDING = 8;
const COMMANDS_PADDING = COMMANDS_OVERFLOW_PADDING + 8;
/** Execute while a query runs ("Executing…" and a spinner) */
const EXECUTING_WIDTH = 124;
const MORE_WIDTH = 32;
/**
 * The least each side needs: Execute, More and Settings, with the one divider before Settings
 * (they never overflow), and one tab with More. At the default 900 px window (480 px builder pane)
 * the row is about 414 px.
 */
const commandsMinWidth = (isExecuting: boolean) =>
	(isExecuting ? EXECUTING_WIDTH : COMMAND_WIDTH.execute) +
	MORE_WIDTH +
	COMMAND_WIDTH.settings +
	DIVIDER_WIDTH +
	COMMANDS_PADDING;
const TABS_MIN_WIDTH = 184;

const useStyles = makeStyles({
	header: {
		display: "flex",
		alignItems: "center",
		borderBottom: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
		paddingRight: tokens.spacingHorizontalS,
		// Below both minimums (a very wide builder pane) the row clips at its end, never overlaps
		overflow: "hidden",
	},
	// Overflow keeps one tab and More however narrow, so the side never shrinks below that
	tabs: { minWidth: `${TABS_MIN_WIDTH}px` },
});

export interface PreviewHeaderProps {
	selectedTab: PreviewTab;
	onTabSelect: (tab: PreviewTab) => void;
	advisoryCount: number;
	execute: {
		onClick: () => void;
		/** A query is running: "Executing…" */
		isExecuting: boolean;
		/** Running, or loading rows: Execute can't be pressed but stays focusable */
		busy: boolean;
		disabledReason?: string;
	};
	/** Save view; Save shows on the FetchXML tab only */
	saveView?: SaveViewCommandOptions;
	/** Send to tool; no command without tools */
	sendToTool?: SendToToolCommandOptions;
	/** Launched by another tool: Return FetchXML shows */
	isCallee?: boolean;
	getCurrentXml: () => string;
	onOpenSettings?: () => void;
}

export function PreviewHeader({
	selectedTab,
	onTabSelect,
	advisoryCount,
	execute,
	saveView,
	sendToTool,
	isCallee,
	getCurrentXml,
	onOpenSettings,
}: PreviewHeaderProps) {
	const styles = useStyles();
	// Hooks run whether or not their command shows, so their state (an open dialog) survives
	const save = useSaveViewCommand(
		saveView ?? {
			fetchXml: "",
			layoutXml: "",
			entityLogicalName: "",
			objectTypeCode: 0,
			primaryIdAttribute: "",
			loadedView: null,
			disabled: true,
		},
	);
	const send = useSendToToolCommand(
		sendToTool ?? { fetchXml: "", entityLogicalName: "", tools: [] },
	);
	const returnCommand = useReturnFetchXmlCommand(getCurrentXml);

	const tabs: BarTab[] = [
		{ value: "xml", label: "FetchXML" },
		{ value: "layout", label: "LayoutXML" },
		{ value: "results", label: "Results" },
		{
			value: "advisories",
			label: "Advisories",
			ariaLabel: `Advisories (${advisoryCount})`,
			content: (
				<>
					{" "}
					<CounterBadge
						count={advisoryCount}
						size="small"
						appearance="ghost"
						color="brand"
					/>
				</>
			),
		},
		{ value: "code", label: "Code", icon: CodeIcon },
	];

	const commands: BarCommand[] = [
		{
			id: "execute",
			label: execute.isExecuting ? "Executing…" : "Execute",
			icon: PlayIcon,
			group: "run",
			pinned: true,
			kind: "button",
			appearance: "primary",
			busy: execute.busy,
			disabledReason: execute.disabledReason,
			onClick: execute.onClick,
		},
	];
	if (isCallee) commands.push(returnCommand);
	if (saveView && selectedTab === "xml") commands.push(save.command);
	if (send) commands.push(send);
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

	const commandsWidth = commands.reduce(
		(sum, c, i) =>
			sum +
			(c.id === "execute" && execute.isExecuting
				? EXECUTING_WIDTH
				: (COMMAND_WIDTH[c.id] ?? 120)) +
			(i > 0 ? DIVIDER_WIDTH : 0),
		COMMANDS_PADDING,
	);

	return (
		<div className={styles.header}>
			<div className={styles.tabs} style={{ flex: `1 1 ${TABS_WIDTH}px` }}>
				<OverflowTabList
					ariaLabel="Query views"
					tabs={tabs}
					selectedValue={selectedTab}
					onTabSelect={(value) => onTabSelect(value as PreviewTab)}
				/>
			</div>
			<div
				style={{
					flex: `1 1 ${commandsWidth}px`,
					minWidth: `${commandsMinWidth(execute.isExecuting)}px`,
				}}
			>
				<OverflowToolbar
					ariaLabel="Query commands"
					size="small"
					align="end"
					// Return and Send to tool can appear after the first layout (launch, tool discovery)
					groups={["run", "return", "file", "share", "settings"]}
					padding={COMMANDS_OVERFLOW_PADDING}
					commands={commands}
				/>
			</div>
			{save.dialog}
		</div>
	);
}
