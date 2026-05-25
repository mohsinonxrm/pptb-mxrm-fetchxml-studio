/**
 * Settings Drawer
 * Contains display preferences for FetchXML Studio
 */

import { useCallback, useState } from "react";
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
	Tooltip,
	Input,
	Field,
} from "@fluentui/react-components";
import {
	Dismiss24Regular,
	Settings20Regular,
	Info16Regular,
	Add20Regular,
	Dismiss16Regular,
	PlugConnected20Regular,
} from "@fluentui/react-icons";
import type {
	DisplaySettings,
	ValueDisplayMode,
	EntityScopeMode,
} from "../../model/displaySettings";
import type { AccessSummary } from "../../api/pptbClient";

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
	radioGroup: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
		marginTop: tokens.spacingVerticalXS,
	},
	disabledHint: {
		fontSize: tokens.fontSizeBase100,
		color: tokens.colorNeutralForegroundDisabled,
		fontStyle: "italic",
	},
	toolInputRow: {
		display: "flex",
		gap: tokens.spacingHorizontalS,
		alignItems: "flex-end",
		marginTop: tokens.spacingVerticalXS,
	},
	toolInput: {
		flex: 1,
	},
	toolList: {
		listStyle: "none",
		margin: 0,
		padding: 0,
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXS,
		marginTop: tokens.spacingVerticalS,
	},
	toolListItem: {
		display: "flex",
		alignItems: "center",
		gap: tokens.spacingHorizontalS,
		background: tokens.colorNeutralBackground3,
		borderRadius: tokens.borderRadiusMedium,
		padding: `${tokens.spacingVerticalXXS} ${tokens.spacingHorizontalS}`,
	},
	toolListItemId: {
		flex: 1,
		fontSize: tokens.fontSizeBase200,
		color: tokens.colorNeutralForeground2,
		overflowWrap: "anywhere" as const,
	},
	toolListItemIcon: {
		color: tokens.colorBrandForeground1,
		flexShrink: 0,
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
	/**
	 * Access summary from useAccessMode — used to constrain which Entity Scope
	 * options are available based on the user's Dataverse privileges.
	 * When null (still loading), all options are shown.
	 */
	accessSummary?: AccessSummary | null;
}

export function SettingsDrawer({
	open,
	settings,
	onClose,
	onSettingsChange,
	accessSummary,
}: SettingsDrawerProps) {
	const styles = useStyles();
	const [newToolId, setNewToolId] = useState("");

	const handleAddTool = useCallback(() => {
		const trimmed = newToolId.trim();
		if (!trimmed || settings.targetTools.includes(trimmed)) return;
		onSettingsChange({ ...settings, targetTools: [...settings.targetTools, trimmed] });
		setNewToolId("");
	}, [newToolId, settings, onSettingsChange]);

	const handleRemoveTool = useCallback(
		(toolId: string) => {
			onSettingsChange({
				...settings,
				targetTools: settings.targetTools.filter((t) => t !== toolId),
			});
		},
		[settings, onSettingsChange],
	);

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

	const handleEntityScopeModeChange = useCallback(
		(mode: EntityScopeMode) => {
			onSettingsChange({ ...settings, entityScopeMode: mode });
		},
		[settings, onSettingsChange],
	);

	const handleAdvancedFindOnlyChange = useCallback(
		(checked: boolean) => {
			onSettingsChange({ ...settings, advancedFindOnly: checked });
		},
		[settings, onSettingsChange],
	);

	// Determine which scope options are available given the user's privilege ceiling
	const canUsePublisherSolution = !accessSummary || accessSummary.fullFilterMode;
	const canUseSolutionOnly =
		!accessSummary || accessSummary.fullFilterMode || accessSummary.solutionsOnlyMode;

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
					<Text className={styles.sectionTitle}>Query Scope</Text>

					{/* Entity Scope Mode */}
					<div className={styles.settingItem}>
						<Text className={styles.settingLabel}>
							Entity Source
							<Tooltip
								content="Controls how the entity list is populated in the toolbar. Scoped modes are faster; 'All Entities' loads every entity in the environment."
								relationship="description"
							>
								<Info16Regular style={{ color: tokens.colorNeutralForeground3 }} />
							</Tooltip>
						</Text>
						<Text className={styles.settingDescription} block>
							How to scope the entity picker
						</Text>
						<RadioGroup
							className={styles.radioGroup}
							value={settings.entityScopeMode}
							onChange={(_e, data) => handleEntityScopeModeChange(data.value as EntityScopeMode)}
						>
							<Tooltip
								content={
									!canUsePublisherSolution ? "Requires prvReadPublisher + prvReadSolution" : ""
								}
								relationship="description"
								positioning="before"
							>
								<Radio
									value="publisher-solution"
									label="Publisher → Solution (default)"
									disabled={!canUsePublisherSolution}
								/>
							</Tooltip>
							<Tooltip
								content={!canUseSolutionOnly ? "Requires prvReadSolution" : ""}
								relationship="description"
								positioning="before"
							>
								<Radio value="solution-only" label="Solution only" disabled={!canUseSolutionOnly} />
							</Tooltip>
							<Radio value="all" label="All Entities" />
						</RadioGroup>
					</div>

					<Divider style={{ marginBottom: tokens.spacingVerticalM }} />

					{/* Advanced Find Only */}
					<div className={styles.settingItem}>
						<div className={styles.settingRow}>
							<div>
								<Text className={styles.settingLabel}>
									Advanced Find Only
									<Tooltip
										content="When on, entities and attributes are limited to those marked IsValidForAdvancedFind — the same set exposed by Advanced Find in Model-Driven Apps. Turn off to access all entities and attributes."
										relationship="description"
									>
										<Info16Regular style={{ color: tokens.colorNeutralForeground3 }} />
									</Tooltip>
								</Text>
								<Text className={styles.settingDescription} block>
									Limit entities &amp; attributes to Advanced Find–eligible ones
								</Text>
								{!settings.advancedFindOnly && (
									<Text className={styles.disabledHint} block>
										All entities and attributes are shown — including non-queryable ones
									</Text>
								)}
							</div>
							<Switch
								checked={settings.advancedFindOnly}
								onChange={(_e, data) => handleAdvancedFindOnlyChange(data.checked)}
							/>
						</div>
					</div>
				</div>

				<Divider />

				{/* ── Display Section ──────────────────────────────────── */}
				<div className={styles.section} style={{ marginTop: tokens.spacingVerticalL }}>
					<Text className={styles.sectionTitle}>Display</Text>

					{/* Logical Names Toggle */}
					<div className={styles.settingItem}>
						<div className={styles.settingRow}>
							<div>
								<Text className={styles.settingLabel}>Use Logical Names</Text>
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
						<Text className={styles.settingLabel}>Value Display Mode</Text>
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
				</div>
				<Divider />

				{/* ── Tool Integration Section ────────────────────────── */}
				<div className={styles.section} style={{ marginTop: tokens.spacingVerticalL }}>
					<Text className={styles.sectionTitle}>
						<PlugConnected20Regular />
						Tool Integration
					</Text>

					<div className={styles.settingItem}>
						<Text className={styles.settingLabel}>
							Target Tools
							<Tooltip
								content="npm package IDs of PPTB tools this studio can send FetchXML queries to via Tool-to-Tool invocation. Example: @linked365/pptb-bulk-data-studio"
								relationship="description"
							>
								<Info16Regular style={{ color: tokens.colorNeutralForeground3 }} />
							</Tooltip>
						</Text>
						<Text className={styles.settingDescription} block>
							Tools that will appear in the &ldquo;Send to Tool&rdquo; button
						</Text>

						<Field label="Tool package ID">
							<div className={styles.toolInputRow}>
								<Input
									className={styles.toolInput}
									placeholder="e.g. @linked365/pptb-bulk-data-studio"
									value={newToolId}
									onChange={(_e, data) => setNewToolId(data.value)}
									onKeyDown={(e) => e.key === "Enter" && handleAddTool()}
								/>
								<Button
									appearance="primary"
									icon={<Add20Regular />}
									onClick={handleAddTool}
									disabled={!newToolId.trim() || settings.targetTools.includes(newToolId.trim())}
									aria-label="Add tool"
								/>
							</div>
						</Field>

						{settings.targetTools.length > 0 && (
							<ul className={styles.toolList}>
								{settings.targetTools.map((toolId) => (
									<li key={toolId} className={styles.toolListItem}>
										<PlugConnected20Regular className={styles.toolListItemIcon} />
										<span className={styles.toolListItemId}>{toolId}</span>
										<Button
											appearance="subtle"
											size="small"
											icon={<Dismiss16Regular />}
											onClick={() => handleRemoveTool(toolId)}
											aria-label={`Remove ${toolId}`}
										/>
									</li>
								))}
							</ul>
						)}
					</div>
				</div>
			</DrawerBody>
		</OverlayDrawer>
	);
}
