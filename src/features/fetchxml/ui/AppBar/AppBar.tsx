/**
 * The app bar (#278): one full-width row above both panes with the query's commands, in the
 * order of a model-driven app's command bar: Run | New · Open view · Save · Send to · Return
 * FetchXML … More | Settings. Run comes first, is the one primary button and never overflows;
 * Settings never does either. What doesn't fit moves into More by priority (D: Save 60, New 40,
 * Open view 30, Send to 20, Return FetchXML 10). More always has Copy FetchXML.
 *
 * New and Open view ask before they discard unsaved changes. Save and Send to keep their editor
 * sync (EditorSyncContext), so AppShell renders this inside the provider. Operations joins in #296.
 */
import { useEffect, useRef, useState, type MutableRefObject } from "react";
import { Badge, Body1Strong, MenuItem, makeStyles, tokens } from "@fluentui/react-components";
import {
	Add20Filled,
	Add20Regular,
	Copy20Regular,
	FolderOpen20Filled,
	Keyboard20Regular,
	FolderOpen20Regular,
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
import { discardedToast, openViewDisabledReason } from "./appBarRules";
import { OpenViewPanel } from "./OpenViewPanel";
import { toLoadedViewInfo, type LoadedViewInfo, type SavedView } from "../../api/pptbClient";

const RunIcon = bundleIcon(Play20Filled, Play20Regular);
const NewIcon = bundleIcon(Add20Filled, Add20Regular);
const OpenViewIcon = bundleIcon(FolderOpen20Filled, FolderOpen20Regular);
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
	/** The query differs from the view as loaded or saved (its FetchXML, columns or widths): the badge */
	unsaved: boolean;
	/** The same, now (the editor's text this moment): New asks first when it's true */
	hasUnsavedChanges?: () => boolean;
	/** Start a new query: the caller resets the tree, the editor and the results */
	onNewQuery: () => void;
	/** Open view: the query's entity and what loading one of its views needs */
	openView: OpenViewOptions;
	/** More › Keyboard shortcuts (#186) */
	onShowShortcuts?: () => void;
	/**
	 * Ctrl+S (#186): set to what selecting Save does, or null while Save can't be pressed, so the
	 * app's key handler presses the same Save
	 */
	saveShortcutRef?: MutableRefObject<(() => void) | null>;
}

export interface OpenViewOptions {
	/** The query's entity; Open view waits for one */
	entityLogicalName?: string;
	entityDisplayName?: string;
	/** From the entity's metadata: Open view waits for it too (running a view by ID needs it) */
	entitySetName?: string;
	/** The view the query is, checked in the list */
	currentViewId?: string;
	/**
	 * Load the view: its FetchXML, layout and identity, and drop the editor's text (AppShell). False
	 * when its FetchXML couldn't be loaded and the query stays as it was.
	 */
	onOpen: (viewInfo: LoadedViewInfo) => boolean | void;
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
	hasUnsavedChanges,
	onNewQuery,
	openView,
	onShowShortcuts,
	saveShortcutRef,
}: AppBarProps) {
	const styles = useStyles();
	const copy = useCopy();
	const { notify } = useNotify();
	// New and Open view with unsaved changes ask first (#278). "Save and start new" / "Save and open"
	// continue once their own save is done, never after a later one.
	const [confirm, setConfirm] = useState<
		{ intent: "new" } | { intent: "open"; view: LoadedViewInfo } | null
	>(null);
	const afterSave = useRef<(() => void) | null>(null);
	// What the dialog last asked about: its wording stays while it closes
	const lastConfirm = useRef<NonNullable<typeof confirm>>({ intent: "new" });
	if (confirm) lastConfirm.current = confirm;
	const asked = lastConfirm.current;
	const [openViewOpen, setOpenViewOpen] = useState(false);
	const save = useSaveViewCommand({
		...saveView,
		onSaveComplete: (...args) => {
			saveView.onSaveComplete?.(...args);
			const next = afterSave.current;
			afterSave.current = null;
			next?.();
		},
		onCancel: () => {
			afterSave.current = null;
			saveView.onCancel?.();
		},
	});
	// Ctrl+S presses Save as it is now (#186)
	useEffect(() => {
		if (!saveShortcutRef) return;
		saveShortcutRef.current = save.command.disabledReason
			? null
			: (save.command.onClick ?? null);
	});
	const isUnsaved = () => (hasUnsavedChanges ? hasUnsavedChanges() : unsaved);
	const handleNew = () => {
		if (isUnsaved()) setConfirm({ intent: "new" });
		else onNewQuery();
	};
	const handlePickView = (view: SavedView) => {
		setOpenViewOpen(false);
		if (!openView.entitySetName) return;
		const viewInfo = toLoadedViewInfo(view, openView.entitySetName);
		if (isUnsaved()) setConfirm({ intent: "open", view: viewInfo });
		else openView.onOpen(viewInfo);
	};
	/** Continue with New or Open view; false when the view couldn't be opened (nothing was replaced) */
	const proceed = (choice: NonNullable<typeof confirm>): boolean => {
		if (choice.intent === "open") return openView.onOpen(choice.view) !== false;
		onNewQuery();
		return true;
	};
	const openViewReason = openViewDisabledReason(
		openView.entityLogicalName,
		openView.entitySetName,
	);
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
		{
			id: "open-view",
			label: "Open view",
			icon: OpenViewIcon,
			group: "file",
			priority: 30,
			kind: "popover",
			hint: "Open a view of this entity",
			disabledReason: openViewReason,
			popoverOpen: openViewOpen,
			onPopoverOpenChange: setOpenViewOpen,
			popover:
				openView.entityLogicalName && !openViewReason ? (
					<OpenViewPanel
						entityLogicalName={openView.entityLogicalName}
						entityDisplayName={openView.entityDisplayName ?? openView.entityLogicalName}
						currentViewId={openView.currentViewId}
						onPick={handlePickView}
					/>
				) : null,
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
					<>
						<MenuItem
							icon={<Copy20Regular />}
							disabled={!getCurrentXml()}
							onClick={() => void copy(getCurrentXml(), "FetchXML")}
						>
							Copy FetchXML
						</MenuItem>
						{onShowShortcuts && (
							<MenuItem
								icon={<Keyboard20Regular />}
								secondaryContent={formatKeys(getEntry("shortcuts"))[0]}
								onClick={onShowShortcuts}
							>
								Keyboard shortcuts
							</MenuItem>
						)}
					</>
				}
			/>
			{save.dialog}
			<DiscardChangesDialog
				open={confirm !== null}
				intent={asked.intent}
				viewName={viewName}
				targetViewName={asked.intent === "open" ? asked.view.name : undefined}
				onSave={() => {
					if (!confirm) return;
					// Save can't run yet (no entity metadata, say): say why, and keep the choice open
					if (save.command.disabledReason) {
						notify.warning({
							title: "Can't save yet",
							body: save.command.disabledReason,
						});
						return;
					}
					setConfirm(null);
					afterSave.current = () => void proceed(confirm);
					save.command.onClick?.();
				}}
				onDiscard={() => {
					if (!confirm) return;
					setConfirm(null);
					// A view that couldn't be loaded replaced nothing: its own error toast says so
					if (proceed(confirm)) notify.info(discardedToast(viewName));
				}}
				onCancel={() => setConfirm(null)}
			/>
		</>
	);
}
