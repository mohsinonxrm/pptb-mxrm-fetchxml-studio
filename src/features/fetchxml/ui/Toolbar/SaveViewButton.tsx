/**
 * SaveViewButton - SplitButton for saving views to Dataverse
 * Supports Save, Save As, and Save and Publish actions
 * Handles both System Views (savedquery) and Personal Views (userquery)
 */

import { forwardRef, useState, useCallback, useEffect, useImperativeHandle } from "react";
import {
	Menu,
	MenuTrigger,
	MenuList,
	MenuItem,
	MenuPopover,
	SplitButton,
	makeStyles,
	tokens,
	Tooltip,
	type MenuButtonProps,
} from "@fluentui/react-components";
import { SaveRegular, SaveArrowRightRegular, SendRegular } from "@fluentui/react-icons";
import {
	SaveViewDialog,
	type SaveViewDialogLoadedView,
	type SaveViewDialogMode,
	type SavedViewDetails,
} from "../Dialogs/SaveViewDialog";
import { useEditorSync } from "../RightPane/editorSync";

const useStyles = makeStyles({
	splitButton: {
		minWidth: "110px",
	},
	menuItem: {
		display: "flex",
		alignItems: "center",
		gap: "8px",
	},
	menuIcon: {
		fontSize: "16px",
		color: tokens.colorNeutralForeground2,
	},
});

interface SaveViewButtonProps {
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
	/** Disabled state */
	disabled?: boolean;
}

/** Save actions for the toolbar's overflow menu (when the Save button doesn't fit) */
export interface SaveViewButtonHandle {
	save: () => void;
	saveAs: () => void;
	saveAndPublish: () => void;
}

export const SaveViewButton = forwardRef<SaveViewButtonHandle, SaveViewButtonProps>(function SaveViewButton({
	fetchXml,
	layoutXml,
	entityLogicalName,
	entityDisplayName,
	objectTypeCode,
	primaryIdAttribute,
	loadedView,
	onSaveComplete,
	disabled = false,
}, ref) {
	const styles = useStyles();
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
			if (synced.status === "cancelled") return;
			if (synced.status === "unchanged") {
				openDialog(mode, publish);
				return;
			}
			// The tree was replaced, which clears the loaded view like any tree edit: save as a new view
			setPendingOpen({ entity: synced.fetchQuery.entity.name, xml: synced.fetchXml });
		},
		[sync, openDialog],
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

	useImperativeHandle(
		ref,
		() => ({ save: handleSave, saveAs: handleSaveAs, saveAndPublish: handleSaveAndPublish }),
		[handleSave, handleSaveAs, handleSaveAndPublish],
	);

	const handleDialogClose = useCallback(() => {
		setDialogOpen(false);
	}, []);

	const handleDialogSaveComplete = useCallback(
		(viewId: string, viewType: "system" | "personal", viewName: string, details: SavedViewDetails) => {
			setDialogOpen(false);
			onSaveComplete?.(viewId, viewType, viewName, details);
		},
		[onSaveComplete]
	);

	// Primary button text based on state
	const primaryText = loadedView ? "Save" : "Save as";
	const primaryIcon = loadedView ? <SaveRegular /> : <SaveArrowRightRegular />;

	return (
		<>
			<Menu positioning="below-end">
				<MenuTrigger disableButtonEnhancement>
					{(triggerProps: MenuButtonProps) => (
						<Tooltip
							content={loadedView ? `Save changes to "${loadedView.name}"` : "Save as new view"}
							relationship="description"
						>
							<SplitButton
								className={styles.splitButton}
								appearance="primary"
								icon={primaryIcon}
								menuButton={triggerProps}
								primaryActionButton={{
									onClick: loadedView ? handleSave : handleSaveAs,
								}}
								disabled={disabled && !editorHasContent}
							>
								{primaryText}
							</SplitButton>
						</Tooltip>
					)}
				</MenuTrigger>
				<MenuPopover>
					<MenuList>
						{loadedView && (
							<MenuItem icon={<SaveRegular className={styles.menuIcon} />} onClick={handleSave}>
								<span className={styles.menuItem}>Save</span>
							</MenuItem>
						)}
						<MenuItem
							icon={<SaveArrowRightRegular className={styles.menuIcon} />}
							onClick={handleSaveAs}
						>
							<span className={styles.menuItem}>Save as…</span>
						</MenuItem>
						{loadedView?.type === "system" && (
							<MenuItem
								icon={<SendRegular className={styles.menuIcon} />}
								onClick={handleSaveAndPublish}
							>
								<span className={styles.menuItem}>Save and publish</span>
							</MenuItem>
						)}
					</MenuList>
				</MenuPopover>
			</Menu>

			<SaveViewDialog
				open={dialogOpen}
				onClose={handleDialogClose}
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
		</>
	);
});
