/**
 * SendToToolButton – launches another PPTB tool with the current FetchXML query as prefill.
 *
 * Reads target tool IDs from displaySettings.targetTools. When no tools are configured,
 * the button is hidden entirely (targets are added under Settings → Tool Integration).
 * When one tool is configured, renders as a single button. When multiple tools are
 * configured, renders a dropdown menu to choose the target.
 */

import { useState, useCallback } from "react";
import {
	Button,
	Menu,
	MenuTrigger,
	MenuList,
	MenuItem,
	MenuPopover,
	Tooltip,
	Spinner,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { PlugConnected20Regular } from "@fluentui/react-icons";
import { sendFetchXmlToTool, type FetchXmlStudioT2TPrefill } from "../../api/invocation";

const useStyles = makeStyles({
	menuItem: {
		display: "flex",
		flexDirection: "column",
		alignItems: "flex-start",
	},
	toolIdHint: {
		fontSize: tokens.fontSizeBase100,
		color: tokens.colorNeutralForeground3,
	},
});

export interface SendToToolButtonProps {
	/** Current serialized FetchXML query */
	fetchXml: string;
	/** Root entity logical name */
	entityLogicalName: string;
	/** List of target tool npm package IDs from settings */
	targetTools: string[];
	/** Disabled state (e.g. no query built yet) */
	disabled?: boolean;
}

/** Extract a short display label from a scoped npm package ID. */
function toolLabel(toolId: string): string {
	// "@linked365/pptb-bulk-data-studio" → "pptb-bulk-data-studio"
	const parts = toolId.split("/");
	return parts[parts.length - 1] ?? toolId;
}

export function SendToToolButton({
	fetchXml,
	entityLogicalName,
	targetTools,
	disabled,
}: SendToToolButtonProps) {
	const styles = useStyles();
	const [isSending, setIsSending] = useState(false);

	const handleSend = useCallback(
		async (targetToolId: string) => {
			if (!fetchXml || !entityLogicalName) return;
			setIsSending(true);
			try {
				const prefill: FetchXmlStudioT2TPrefill = { fetchXml, entityLogicalName };
				await sendFetchXmlToTool(targetToolId, prefill);
			} catch (err) {
				const message = err instanceof Error ? err.message : String(err);
				await window.toolboxAPI?.utils?.showNotification?.({
					title: "Cannot open tool",
					body: message,
					type: "error",
				});
			} finally {
				setIsSending(false);
			}
		},
		[fetchXml, entityLogicalName],
	);

	const isDisabled = disabled || isSending || !fetchXml || !entityLogicalName;
	const icon = isSending ? <Spinner size="tiny" /> : <PlugConnected20Regular />;

	// No tools configured → nothing to send to. The button is hidden entirely;
	// configuration lives under Settings → Tool Integration. (AppShell also gates
	// on this, so this is a defensive guard.)
	if (targetTools.length === 0) {
		return null;
	}

	// Single tool → direct button
	if (targetTools.length === 1) {
		const toolId = targetTools[0];
		return (
			<Tooltip content={`Send FetchXML to ${toolId}`} relationship="description">
				<Button
					appearance="subtle"
					icon={icon}
					disabled={isDisabled}
					onClick={() => handleSend(toolId)}
				>
					{toolLabel(toolId)}
				</Button>
			</Tooltip>
		);
	}

	// Multiple tools → dropdown menu
	return (
		<Menu>
			<MenuTrigger disableButtonEnhancement>
				<Button appearance="subtle" icon={icon} disabled={isDisabled}>
					Send to Tool
				</Button>
			</MenuTrigger>
			<MenuPopover>
				<MenuList>
					{targetTools.map((toolId) => (
						<MenuItem
							key={toolId}
							icon={<PlugConnected20Regular />}
							className={styles.menuItem}
							onClick={() => handleSend(toolId)}
						>
							{toolLabel(toolId)}
							<span className={styles.toolIdHint}>{toolId}</span>
						</MenuItem>
					))}
				</MenuList>
			</MenuPopover>
		</Menu>
	);
}
