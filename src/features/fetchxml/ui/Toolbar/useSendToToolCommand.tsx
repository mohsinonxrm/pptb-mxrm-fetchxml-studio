/**
 * Send to tool as a command for the shared bar (#232): launches another PPTB tool with the current
 * FetchXML query as prefill.
 *
 * Targets are discovered automatically via the host capability registry (tools that declare
 * the "fetchxml" capability); see useSendToTools. It's always a "Send to tool" menu listing the
 * discovered tools (even a single one), with each tool's runtime id under its name. With no tools
 * there's no command (null).
 *
 * The launch is fire-and-forget: launchTool's promise doesn't resolve until the callee window
 * closes, so we deliberately do NOT await it (that would pin the command in a loading state for
 * the whole session). We pass noReturn and never consume a result.
 */

import { useCallback } from "react";
import { MenuItem, MenuList, makeStyles, tokens } from "@fluentui/react-components";
import { PlugConnected20Filled, PlugConnected20Regular, bundleIcon } from "@fluentui/react-icons";
import { useNotify } from "../../../../shared/feedback/useNotify";
import {
	sendFetchXmlToTool,
	type DiscoveredTool,
	type FetchXmlStudioT2TPrefill,
} from "../../api/invocation";
import { useEditorSync } from "../RightPane/editorSync";
import type { BarCommand } from "../../../../shared/components/commandBar";
import { NO_QUERY_REASON } from "./useSaveViewCommand";

const SendIcon = bundleIcon(PlugConnected20Filled, PlugConnected20Regular);

const useStyles = makeStyles({
	menuItemContent: {
		display: "flex",
		flexDirection: "column",
		alignItems: "flex-start",
		gap: tokens.spacingVerticalXXS,
	},
	toolIdHint: {
		fontSize: tokens.fontSizeBase100,
		color: tokens.colorNeutralForeground3,
	},
});

export interface SendToToolCommandOptions {
	/** Current serialized FetchXML query */
	fetchXml: string;
	/** Root entity logical name */
	entityLogicalName: string;
	/** Tools discovered via the "fetchxml" capability */
	tools: DiscoveredTool[];
	/** No query to send yet */
	disabled?: boolean;
}

export function useSendToToolCommand({
	fetchXml,
	entityLogicalName,
	tools,
	disabled,
}: SendToToolCommandOptions): BarCommand | null {
	const styles = useStyles();
	const { sync, editorHasContent } = useEditorSync();
	const { notify } = useNotify();

	const handleSend = useCallback(
		async (targetToolId: string) => {
			// Editor mode: load the editor's FetchXML into the tree first, then send that (#102)
			const synced = await sync("send");
			if (synced.status === "cancelled") return;
			const xmlToSend = synced.status === "synced" ? synced.fetchXml : fetchXml;
			const entityToSend =
				synced.status === "synced" ? synced.fetchQuery.entity.name : entityLogicalName;
			if (!xmlToSend || !entityToSend) return;
			const prefill: FetchXmlStudioT2TPrefill = {
				fetchXml: xmlToSend,
				entityLogicalName: entityToSend,
			};
			// Fire-and-forget — do not await (see file header). Errors are surfaced async.
			void sendFetchXmlToTool(targetToolId, prefill).catch(async (err) => {
				const message = err instanceof Error ? err.message : String(err);
				// A rapid second launch is rejected by the host's one-at-a-time guard;
				// that's benign, so don't nag the user about it.
				if (/already in progress/i.test(message)) return;
				// In the app's Toaster: it's about something the user started here (#277)
				const tool = tools.find((t) => t.id === targetToolId)?.name ?? "the tool";
				notify.error({ title: `Couldn't open ${tool}`, body: message });
			});
		},
		[fetchXml, entityLogicalName, sync, tools, notify],
	);

	// No tools discovered → nothing to send to
	if (tools.length === 0) return null;

	// With content in the editor, sending is available even before the tree has a query
	const isDisabled = !editorHasContent && (disabled || !fetchXml || !entityLogicalName);

	return {
		id: "send",
		label: "Send to tool",
		icon: SendIcon,
		group: "share",
		priority: 20,
		kind: "menu",
		disabledReason: isDisabled ? NO_QUERY_REASON : undefined,
		menu: (
			<MenuList>
				{tools.map((tool) => (
					<MenuItem
						key={tool.id}
						icon={<PlugConnected20Regular />}
						onClick={() => void handleSend(tool.id)}
					>
						<div className={styles.menuItemContent}>
							<span>{tool.name}</span>
							<span className={styles.toolIdHint}>{tool.id}</span>
						</div>
					</MenuItem>
				))}
			</MenuList>
		),
	};
}
