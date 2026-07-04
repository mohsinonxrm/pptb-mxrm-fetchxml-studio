/**
 * SendToToolButton – launches another PPTB tool with the current FetchXML query as prefill.
 *
 * Targets are discovered automatically via the host capability registry (tools that declare
 * the "fetchxml" capability); see useSendToTools. Always renders as a "Send to Tool" dropdown
 * listing the discovered tools (even a single one), with each tool's runtime id shown beneath
 * its name. When no tools are found the button is hidden entirely (AppShell gates on this).
 *
 * The launch is fire-and-forget: launchTool's promise doesn't resolve until the callee window
 * closes, so we deliberately do NOT await it (that would pin the button in a loading state for
 * the whole session). We pass noReturn and never consume a result.
 */

import { useCallback } from "react";
import {
	Button,
	Menu,
	MenuTrigger,
	MenuList,
	MenuItem,
	MenuPopover,
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
	menuItemContent: {
		display: "flex",
		flexDirection: "column",
		alignItems: "flex-start",
		gap: "2px",
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

	const handleSend = useCallback(
		(targetToolId: string) => {
			if (!fetchXml || !entityLogicalName) return;
			const prefill: FetchXmlStudioT2TPrefill = { fetchXml, entityLogicalName };
			// Fire-and-forget — do not await (see file header). Errors are surfaced async.
			void sendFetchXmlToTool(targetToolId, prefill).catch(async (err) => {
				const message = err instanceof Error ? err.message : String(err);
				// A rapid second launch is rejected by the host's one-at-a-time guard;
				// that's benign, so don't nag the user about it.
				if (/already in progress/i.test(message)) return;
				await window.toolboxAPI?.utils?.showNotification?.({
					title: "Cannot open tool",
					body: message,
					type: "error",
				});
			});
		},
		[fetchXml, entityLogicalName],
	);

	const isDisabled = disabled || !fetchXml || !entityLogicalName;

	// No tools discovered → nothing to send to. The button is hidden entirely
	// (AppShell also gates on this, so this is a defensive guard).
	if (tools.length === 0) {
		return null;
	}

	return (
		<Menu>
			<MenuTrigger disableButtonEnhancement>
				<Button appearance="subtle" icon={<PlugConnected20Regular />} disabled={isDisabled}>
					Send to Tool
				</Button>
			</MenuTrigger>
			<MenuPopover>
				<MenuList>
					{tools.map((tool) => (
						<MenuItem
							key={tool.id}
							icon={<PlugConnected20Regular />}
							onClick={() => handleSend(tool.id)}
						>
							<div className={styles.menuItemContent}>
								<span>{tool.name}</span>
								<span className={styles.toolIdHint}>{tool.id}</span>
							</div>
						</MenuItem>
					))}
				</MenuList>
			</MenuPopover>
		</Menu>
	);
}
