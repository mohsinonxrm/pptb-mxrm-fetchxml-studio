/**
 * The right pane header's rules (#232), apart from the component so it stays fast-refreshable.
 */
import { NO_QUERY_REASON } from "../Toolbar/useSaveViewCommand";

export type PreviewTab = "xml" | "layout" | "results" | "advisories" | "code";

/** Why Execute is off with an empty editor in editor mode */
export const EMPTY_EDITOR_REASON = "Enter FetchXML first";

/** Execute's reason, if it can't run: the tree's or the editor's query is empty */
export function emptyQueryReason(
	isEditorModeActive: boolean,
	editorHasContent: boolean,
	xml: string,
): string | undefined {
	if (isEditorModeActive) return editorHasContent ? undefined : EMPTY_EDITOR_REASON;
	return xml ? undefined : NO_QUERY_REASON;
}
