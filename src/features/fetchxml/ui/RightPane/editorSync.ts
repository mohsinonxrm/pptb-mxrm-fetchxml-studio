/**
 * Editor-mode sync (#102): in editor mode, actions that work on the query (Execute, Save, Send to
 * Tool) first load the editor's FetchXML into the tree, so the tree is always what runs, what is
 * saved and what is sent. PreviewTabs owns the editor state and provides the sync function;
 * toolbar buttons rendered inside it call it before acting.
 */
import { createContext, useContext } from "react";
import type { FetchNode } from "../../model/nodes";

export type EditorSyncAction = "execute" | "save" | "send";

export type EditorSyncResult =
	/** Not in editor mode, or the editor already matches the tree: use the tree as it is */
	| { status: "unchanged" }
	/** The editor's FetchXML was loaded into the tree; use these (state updates arrive later) */
	| { status: "synced"; fetchXml: string; fetchQuery: FetchNode }
	/** Invalid FetchXML (the error is shown) or the user cancelled: don't continue */
	| { status: "cancelled" };

export interface EditorSyncValue {
	sync: (action: EditorSyncAction) => Promise<EditorSyncResult>;
	/** Editor mode is on and the editor has content: actions are available even with an empty tree */
	editorHasContent: boolean;
}

export const EditorSyncContext = createContext<EditorSyncValue | null>(null);

const NO_SYNC: EditorSyncValue = { sync: async () => ({ status: "unchanged" }), editorHasContent: false };

/** Editor sync for the surrounding PreviewTabs, or a no-op outside it. */
export function useEditorSync(): EditorSyncValue {
	return useContext(EditorSyncContext) ?? NO_SYNC;
}
