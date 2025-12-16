/**
 * SaveViewButton - SplitButton for saving views to Dataverse
 * Supports Save, Save As, and Save and Publish actions
 * Handles both System Views (savedquery) and Personal Views (userquery)
 */

import { useState, useCallback } from "react";
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
import { SaveViewDialog, type SaveViewDialogMode } from "../Dialogs/SaveViewDialog";

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
	/** Entity object type code for LayoutXML */
	objectTypeCode: number;
	/** Primary ID attribute for LayoutXML */
	primaryIdAttribute: string;
	/** Loaded view info (for Save/overwrite) */
	loadedView: {
		id: string;
		type: "system" | "personal";
		name: string;
	} | null;
	/** Callback after successful save */
	onSaveComplete?: (viewId: string, viewType: "system" | "personal", viewName: string) => void;
	/** Disabled state */
	disabled?: boolean;
}

export function SaveViewButton({
	fetchXml,
	layoutXml,
	entityLogicalName,
	objectTypeCode,
	primaryIdAttribute,
	loadedView,
	onSaveComplete,
	disabled = false,
}: SaveViewButtonProps) {
	const styles = useStyles();
	const [dialogOpen, setDialogOpen] = useState(false);
	const [dialogMode, setDialogMode] = useState<SaveViewDialogMode>("save");
	const [shouldPublish, setShouldPublish] = useState(false);

	const handleSave = useCallback(() => {
		setShouldPublish(false);
		setDialogMode("save");
		setDialogOpen(true);
	}, []);

	const handleSaveAs = useCallback(() => {
		setShouldPublish(false);
		setDialogMode("saveAs");
		setDialogOpen(true);
	}, []);

	const handleSaveAndPublish = useCallback(() => {
		setShouldPublish(true);
		setDialogMode("save");
		setDialogOpen(true);
	}, []);

	const handleDialogClose = useCallback(() => {
		setDialogOpen(false);
	}, []);

	const handleDialogSaveComplete = useCallback(
		(viewId: string, viewType: "system" | "personal", viewName: string) => {
			setDialogOpen(false);
			onSaveComplete?.(viewId, viewType, viewName);
		},
		[onSaveComplete]
	);

	// Primary button text based on state
	const primaryText = loadedView ? "Save" : "Save As";
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
								disabled={disabled}
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
							<span className={styles.menuItem}>Save As...</span>
						</MenuItem>
						{loadedView?.type === "system" && (
							<MenuItem
								icon={<SendRegular className={styles.menuIcon} />}
								onClick={handleSaveAndPublish}
							>
								<span className={styles.menuItem}>Save and Publish</span>
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
				objectTypeCode={objectTypeCode}
				primaryIdAttribute={primaryIdAttribute}
				loadedView={loadedView}
			/>
		</>
	);
}
