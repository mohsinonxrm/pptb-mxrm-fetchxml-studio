/**
 * SendToToolButton – launches another PPTB tool with the current FetchXML query as prefill.
 *
 * Targets are discovered automatically via the host capability registry (tools that declare
 * the "fetchxml" capability); see useSendToTools. When one tool is found, renders as a single
 * button; when several are found, renders a dropdown menu to choose the target. When none are
 * found the button is hidden entirely (AppShell gates on this).
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
import {
	sendFetchXmlToTool,
	type DiscoveredTool,
	type FetchXmlStudioT2TPrefill,
} from "../../api/invocation";

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
	/** Tools discovered via the "fetchxml" capability */
	tools: DiscoveredTool[];
	/** Disabled state (e.g. no query built yet) */
	disabled?: boolean;
}

export function SendToToolButton({
	fetchXml,
	entityLogicalName,
	tools,
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

	// No tools discovered → nothing to send to. The button is hidden entirely
	// (AppShell also gates on this, so this is a defensive guard).
	if (tools.length === 0) {
		return null;
	}

	// Single tool → direct button
	if (tools.length === 1) {
		const tool = tools[0];
		return (
			<Tooltip content={`Send FetchXML to ${tool.name}`} relationship="description">
				<Button
					appearance="subtle"
					icon={icon}
					disabled={isDisabled}
					onClick={() => handleSend(tool.id)}
				>
					{tool.name}
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
					{tools.map((tool) => (
						<MenuItem
							key={tool.id}
							icon={<PlugConnected20Regular />}
							className={styles.menuItem}
							onClick={() => handleSend(tool.id)}
						>
							{tool.name}
							<span className={styles.toolIdHint}>{tool.id}</span>
						</MenuItem>
					))}
				</MenuList>
			</MenuPopover>
		</Menu>
	);
}
