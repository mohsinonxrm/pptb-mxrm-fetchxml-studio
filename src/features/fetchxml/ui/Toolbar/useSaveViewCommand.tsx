/**
 * Save view as a command for the shared bar (#232): a split button for Save (or Save as), with
 * Save, Save as… and Save and publish in its menu, plus the dialog it opens. Handles both system
 * views (savedquery) and personal views (userquery).
 *
 * It's a hook rather than a component so the bar can render it on the bar or, once it overflows, in
 * More, while the dialog state lives here once.
 */

import { useState, useCallback, useEffect, type ReactNode } from "react";
import { MenuItem, MenuList } from "@fluentui/react-components";
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
	/** Loaded view info (for Save/overwrite) */
	loadedView: SaveViewDialogLoadedView | null;
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
	loadedView,
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
	const [pendingOpen, setPendingOpen] = useState<{ entity: string; xml: string } | null>(null);

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
			// The tree was replaced, which clears the loaded view like any tree edit: save as a new view
			setPendingOpen({ entity: synced.fetchQuery.entity.name, xml: synced.fetchXml });
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
		openDialog("saveAs", false);
	}, [pendingOpen, fetchXml, entityLogicalName, objectTypeCode, primaryIdAttribute, openDialog]);

	const handleSave = useCallback(() => void requestOpen("save", false), [requestOpen]);
	const handleSaveAs = useCallback(() => void requestOpen("saveAs", false), [requestOpen]);
	const handleSaveAndPublish = useCallback(() => void requestOpen("save", true), [requestOpen]);

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

	const command: BarCommand = {
		id: "save",
		label: loadedView ? "Save" : "Save as",
		icon: loadedView ? SaveIcon : SaveAsIcon,
		group: "file",
		priority: 60,
		kind: "split",
		// One primary per bar: Execute (#232)
		appearance: "secondary",
		onClick: loadedView ? handleSave : handleSaveAs,
		disabledReason,
		hint: loadedView ? `Save changes to "${loadedView.name}"` : "Save as new view",
		// The menu lists Save itself, so More shows it as it is
		menuHasPrimary: true,
		menu: (
			<MenuList>
				{loadedView && (
					<MenuItem icon={<SaveRegular />} onClick={handleSave}>
						Save
					</MenuItem>
				)}
				<MenuItem icon={<SaveArrowRightRegular />} onClick={handleSaveAs}>
					Save as…
				</MenuItem>
				{loadedView?.type === "system" && (
					<MenuItem icon={<SendRegular />} onClick={handleSaveAndPublish}>
						Save and publish
					</MenuItem>
				)}
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
			loadedView={loadedView}
		/>
	);

	return { command, dialog };
}
