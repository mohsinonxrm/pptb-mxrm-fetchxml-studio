/**
 * Settings Drawer
 * Contains display preferences for FetchXML Studio
 */

import { useCallback } from "react";
import {
	DrawerBody,
	DrawerHeader,
	DrawerHeaderTitle,
	OverlayDrawer,
	Button,
	makeStyles,
	tokens,
	Text,
	Switch,
	Dropdown,
	Option,
	Divider,
	Radio,
	RadioGroup,
} from "@fluentui/react-components";
import { Dismiss24Regular, Settings20Regular } from "@fluentui/react-icons";
import type {
	DisplaySettings,
	ValueDisplayMode,
} from "../../model/displaySettings";
import type { TreeIconStyle } from "../../model/treeIcons";

const useStyles = makeStyles({
	drawer: {
		width: "360px",
	},
	section: {
		marginBottom: tokens.spacingVerticalL,
	},
	sectionTitle: {
		fontSize: tokens.fontSizeBase400,
		fontWeight: tokens.fontWeightSemibold,
		marginBottom: tokens.spacingVerticalM,
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
	},
	settingItem: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
		marginBottom: tokens.spacingVerticalM,
	},
	settingRow: {
		display: "flex",
		justifyContent: "space-between",
		alignItems: "center",
	},
	settingLabel: {
		fontWeight: tokens.fontWeightSemibold,
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalXS,
	},
	settingDescription: {
		fontSize: tokens.fontSizeBase200,
		color: tokens.colorNeutralForeground3,
	},
	dropdown: {
		minWidth: "160px",
	},
});

export interface SettingsDrawerProps {
	/** Whether the drawer is open */
	open: boolean;
	/** Current display settings */
	settings: DisplaySettings;
	/** Called when drawer should close */
	onClose: () => void;
	/** Called when settings change */
	onSettingsChange: (settings: DisplaySettings) => void;
}

export function SettingsDrawer({
	open,
	settings,
	onClose,
	onSettingsChange,
}: SettingsDrawerProps) {
	const styles = useStyles();

	const handleLogicalNamesChange = useCallback(
		(checked: boolean) => {
			onSettingsChange({ ...settings, useLogicalNames: checked });
		},
		[settings, onSettingsChange],
	);

	const handleValueDisplayModeChange = useCallback(
		(mode: ValueDisplayMode) => {
			onSettingsChange({ ...settings, valueDisplayMode: mode });
		},
		[settings, onSettingsChange],
	);

	const handleTreeIconStyleChange = useCallback(
		(treeIconStyle: TreeIconStyle) => {
			onSettingsChange({ ...settings, treeIconStyle });
		},
		[settings, onSettingsChange],
	);

	return (
		<OverlayDrawer
			open={open}
			onOpenChange={(_e, data) => !data.open && onClose()}
			position="end"
			size="small"
			className={styles.drawer}
		>
			<DrawerHeader>
				<DrawerHeaderTitle
					action={
						<Button
							appearance="subtle"
							aria-label="Close"
							icon={<Dismiss24Regular />}
							onClick={onClose}
						/>
					}
				>
					<Settings20Regular style={{ marginRight: tokens.spacingHorizontalS }} />
					Settings
				</DrawerHeaderTitle>
			</DrawerHeader>
			<DrawerBody>
				{/* ── Query Scope Section ──────────────────────────────── */}
				<div className={styles.section}>
					<Text className={styles.sectionTitle}>Query scope</Text>

					<Text className={styles.settingDescription} block>
						Set where entities come from (publisher, solutions, all entities) and Advanced Find only from the
						scope button above the entity picker.
					</Text>
				</div>

				<Divider />

				{/* ── Display Section ──────────────────────────────────── */}
				<div className={styles.section} style={{ marginTop: tokens.spacingVerticalL }}>
					<Text className={styles.sectionTitle}>Display</Text>

					{/* Logical Names Toggle */}
					<div className={styles.settingItem}>
						<div className={styles.settingRow}>
							<div>
								<Text className={styles.settingLabel}>Use logical names</Text>
								<Text className={styles.settingDescription} block>
									Show attribute logical names in column headers instead of display names
								</Text>
							</div>
							<Switch
								checked={settings.useLogicalNames}
								onChange={(_e, data) => handleLogicalNamesChange(data.checked)}
							/>
						</div>
					</div>

					{/* Value Display Mode */}
					<div className={styles.settingItem}>
						<Text className={styles.settingLabel}>Value display mode</Text>
						<Text className={styles.settingDescription} block>
							How to display cell values in the results grid
						</Text>
						<Dropdown
							className={styles.dropdown}
							value={
								settings.valueDisplayMode === "formatted"
									? "Formatted"
									: settings.valueDisplayMode === "raw"
										? "Raw"
										: "Both"
							}
							selectedOptions={[settings.valueDisplayMode]}
							onOptionSelect={(_e, data) =>
								handleValueDisplayModeChange(data.optionValue as ValueDisplayMode)
							}
						>
							<Option value="formatted">Formatted</Option>
							<Option value="raw">Raw</Option>
							<Option value="both">Both (2 columns per attribute)</Option>
						</Dropdown>
					</div>

					{/* Tree icons (#230) */}
					<div className={styles.settingItem}>
						<Text className={styles.settingLabel}>Tree icons</Text>
						<Text className={styles.settingDescription} block>
							Each node type has its own icon, coloured by what it does: tables, columns, criteria, sort
						</Text>
						<RadioGroup
							value={settings.treeIconStyle}
							onChange={(_e, data) => handleTreeIconStyleChange(data.value as TreeIconStyle)}
							aria-label="Tree icons"
						>
							<Radio value="tinted" label="Tinted icons" />
							<Radio value="color" label="Color icons (multicolour where available)" />
						</RadioGroup>
					</div>
				</div>
			</DrawerBody>
		</OverlayDrawer>
	);
}
