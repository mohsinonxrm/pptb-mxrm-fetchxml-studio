/**
 * Save view as a command for the shared bar (#232): a split button for Save (or Save as), with
 * Save, Save as… and Save and publish in its menu, plus the dialog it opens. Handles both system
 * views (savedquery) and personal views (userquery).
 *
 * It's a hook rather than a component so the bar can render it on the bar or, once it overflows, in
 * More, while the dialog state lives here once.
 */

import { useState, useCallback, useEffect, type ReactNode } from "react";
import { MenuItem, MenuList, Tooltip } from "@fluentui/react-components";
import {
	Save20Filled,
	Save20Regular,
	SaveArrowRight20Filled,
	SaveArrowRight20Regular,
	SaveRegular,
	SaveArrowRightRegular,
	SendRegular,
	bundleIcon,
} from "@fluentui/react-icons";
import {
	SaveViewDialog,
	type SaveViewDialogLoadedView,
	type SaveViewDialogMode,
	type SavedViewDetails,
} from "../Dialogs/SaveViewDialog";
import { useEditorSync } from "../RightPane/editorSync";
import type { BarCommand } from "../../../../shared/components/commandBar";
import type { SaveAction } from "../../model/saveTarget";
import { useSaveOptions } from "./useSaveOptions";

const SaveIcon = bundleIcon(Save20Filled, Save20Regular);
const SaveAsIcon = bundleIcon(SaveArrowRight20Filled, SaveArrowRight20Regular);

/** Why Save and Send to tool are off before there's a query */
export const NO_QUERY_REASON = "Choose an entity first";

export interface SaveViewCommandOptions {
	/** Current FetchXML to save */
	fetchXml: string;
	/** Current LayoutXML configuration */
	layoutXml: string;
	/** Entity logical name for the view */
	entityLogicalName: string;
	/** Entity display name, for the default view hint (#142) */
	entityDisplayName?: string;
	/** Entity object type code for LayoutXML */
	objectTypeCode: number;
	/** Primary ID attribute for LayoutXML */
	primaryIdAttribute: string;
	/**
	 * The view Save would overwrite: the view the query is (the store's viewIdentity), kept across
	 * edits (#278 PR 3). null for a query that isn't a view.
	 */
	targetView: SaveViewDialogLoadedView | null;
	/** Callback after successful save */
	onSaveComplete?: (
		viewId: string,
		viewType: "system" | "personal",
		viewName: string,
		details: SavedViewDetails,
	) => void;
	/** No query to save yet */
	disabled?: boolean;
	/** The save dialog closed without saving */
	onCancel?: () => void;
}

export function useSaveViewCommand({
	fetchXml,
	layoutXml,
	entityLogicalName,
	entityDisplayName,
	objectTypeCode,
	primaryIdAttribute,
	targetView,
	onSaveComplete,
	disabled = false,
	onCancel,
}: SaveViewCommandOptions): { command: BarCommand; dialog: ReactNode } {
	const [dialogOpen, setDialogOpen] = useState(false);
	const [dialogMode, setDialogMode] = useState<SaveViewDialogMode>("save");
	const [shouldPublish, setShouldPublish] = useState(false);

	const { sync, editorHasContent } = useEditorSync();
	// Editor FetchXML just loaded into the tree (#102): open the dialog once the new tree and its
	// root entity metadata have reached these props, so the view is saved with the right entity.
	const [pendingOpen, setPendingOpen] = useState<{
		entity: string;
		xml: string;
		mode: SaveViewDialogMode;
		publish: boolean;
	} | null>(null);
	// What Save offers for the view, with its checks (owner decision Q3)
	const options = useSaveOptions(targetView);

	const openDialog = useCallback((mode: SaveViewDialogMode, publish: boolean) => {
		setShouldPublish(publish);
		setDialogMode(mode);
		setDialogOpen(true);
	}, []);

	const requestOpen = useCallback(
		async (mode: SaveViewDialogMode, publish: boolean) => {
			const synced = await sync("save");
			// Invalid FetchXML, or the user said no to the changes: nothing opens, so it's a cancel
			if (synced.status === "cancelled") {
				onCancel?.();
				return;
			}
			if (synced.status === "unchanged") {
				openDialog(mode, publish);
				return;
			}
			// The editor's FetchXML replaced the tree: open once the new tree has reached these props
			setPendingOpen({
				entity: synced.fetchQuery.entity.name,
				xml: synced.fetchXml,
				mode,
				publish,
			});
		},
		[sync, openDialog, onCancel],
	);

	useEffect(() => {
		if (!pendingOpen) return;
		const ready =
			fetchXml === pendingOpen.xml &&
			entityLogicalName === pendingOpen.entity &&
			objectTypeCode > 0 &&
			primaryIdAttribute !== "";
		if (!ready) return;
		setPendingOpen(null);
		// Another entity's FetchXML isn't the view any more (its identity went): save it as a new view
		if (pendingOpen.mode === "save" && targetView) openDialog("save", pendingOpen.publish);
		else openDialog("saveAs", false);
	}, [
		pendingOpen,
		fetchXml,
		entityLogicalName,
		objectTypeCode,
		primaryIdAttribute,
		openDialog,
		targetView,
	]);

	const run = (action: SaveAction) =>
		void requestOpen(action.kind === "save-as" ? "saveAs" : "save", action.publishes);
	const iconFor = (action: SaveAction) =>
		action.kind === "save-as" ? (
			<SaveArrowRightRegular />
		) : action.publishes ? (
			<SendRegular />
		) : (
			<SaveRegular />
		);
	const describe = (action: SaveAction) =>
		action.warning ? `${action.label}. ${action.warning}` : action.label;

	const handleDialogSaveComplete = useCallback(
		(
			viewId: string,
			viewType: "system" | "personal",
			viewName: string,
			details: SavedViewDetails,
		) => {
			setDialogOpen(false);
			onSaveComplete?.(viewId, viewType, viewName, details);
		},
		[onSaveComplete],
	);

	// With content in the editor, saving is available even before the tree has a query (#102)
	const disabledReason = disabled && !editorHasContent ? NO_QUERY_REASON : undefined;

	// The button says Save; what it does follows the rules (Q3): save changes to a personal view you
	// can write, or save as a new view. The menu adds the rest, such as overwriting a system view.
	const command: BarCommand = {
		id: "save",
		label: "Save",
		icon: options.primary.kind === "save-as" ? SaveAsIcon : SaveIcon,
		group: "file",
		priority: 60,
		kind: "split",
		// One primary per bar: Run (#232, #278)
		appearance: "secondary",
		onClick: () => run(options.primary),
		disabledReason,
		hint: describe(options.primary),
		// The menu lists the primary action itself, so More shows it as it is
		menuHasPrimary: true,
		menu: (
			<MenuList>
				{[options.primary, ...options.menu].map((action) => {
					const item = (
						<MenuItem
							key={action.kind}
							icon={iconFor(action)}
							onClick={() => run(action)}
						>
							{action.label}
						</MenuItem>
					);
					return action.warning ? (
						<Tooltip
							key={action.kind}
							content={action.warning}
							relationship="description"
							positioning="after"
						>
							{item}
						</Tooltip>
					) : (
						item
					);
				})}
			</MenuList>
		),
	};

	const dialog = (
		<SaveViewDialog
			open={dialogOpen}
			onClose={() => {
				setDialogOpen(false);
				onCancel?.();
			}}
			onSaveComplete={handleDialogSaveComplete}
			mode={dialogMode}
			shouldPublish={shouldPublish}
			fetchXml={fetchXml}
			layoutXml={layoutXml}
			entityLogicalName={entityLogicalName}
			entityDisplayName={entityDisplayName}
			objectTypeCode={objectTypeCode}
			primaryIdAttribute={primaryIdAttribute}
			loadedView={targetView}
		/>
	);

	return { command, dialog };
}
