/**
 * What Save offers for the loaded view (#278, owner decision Q13). Pure, so every rule is unit
 * tested; the app bar's Save split button renders `primary` as its button and `menu` as its menu.
 *
 * The rules follow model-driven apps:
 *  - "Save changes to current view" needs Write on the view; without it, system or personal, you can
 *    only "Save as new view".
 *    https://learn.microsoft.com/power-apps/user/grid-filters-advanced#save-your-view
 *  - "Save changes to current view" is only shown for personal views: a system view can't be updated
 *    from the app.
 *    https://learn.microsoft.com/troubleshoot/power-platform/power-apps/create-and-use-apps/troubleshoot-model-driven-app-view-issues#the-save-changes-to-current-view-option-is-missing
 * FetchXML Studio adds one maker option to that: a customizer can overwrite a customizable system
 * view from the menu, after a confirmation, and the change is published.
 */

import type { SavedView, checkSavedQueryPrivileges } from "../api/pptbClient";
import { canOverwriteView, isSharedView } from "./savedViews";

export type SaveActionKind = "overwrite-personal" | "overwrite-system" | "save-as";

export interface SaveAction {
	kind: SaveActionKind;
	/** The command's label, sentence case; "…" when it opens a dialog first */
	label: string;
	/** Ask before writing: overwriting a system view, or a personal view someone else owns */
	needsConfirm: boolean;
	/** Publish the Entity's customizations after the save (system views only) */
	publishes: boolean;
}

export interface SaveOptions {
	primary: SaveAction;
	menu: SaveAction[];
}

/** The loaded view, as Load view reads it (pptbClient SavedView / builderStore LoadedViewState) */
export type SaveTargetView = Pick<SavedView, "type" | "name" | "isCustomizable" | "ownerId">;

/** What checkSavedQueryPrivileges reports: prvWriteQuery + prvWriteCustomization, and prvPublishCustomization */
export type SystemViewPrivileges = Awaited<ReturnType<typeof checkSavedQueryPrivileges>>;

export interface SaveOptionsInput {
	/** The loaded view (the store's view identity); null when no view is loaded */
	view: SaveTargetView | null;
	/** The connected user (WhoAmI), to tell a personal view shared with the user from their own (#142) */
	currentUserId?: string;
	/**
	 * Personal views someone else owns: whether the user has Write on it (checkUserQueryWriteAccess,
	 * which reads RetrievePrincipalAccess with hasWriteAccess). Undefined when it isn't known yet or
	 * couldn't be read: overwrite is then offered and Dataverse decides (#142).
	 */
	personalWriteAccess?: boolean;
	/** System views: the user's privileges (checkSavedQueryPrivileges); null or undefined while unknown */
	systemViewPrivileges?: SystemViewPrivileges | null;
}

/** A view name inside a label, quoted as in the owner decisions and today's Save tooltip */
const quoted = (name: string): string => `"${name}"`;

export const SAVE_AS_LABEL = "Save as new view…";

const SAVE_AS: SaveAction = { kind: "save-as", label: SAVE_AS_LABEL, needsConfirm: false, publishes: false };

/**
 * Whether the user can write the personal view: their own view always (#142); a view someone else
 * owns unless RetrievePrincipalAccess said it has no WriteAccess. Unknown owners count as the
 * user's own, as in SaveViewDialog: only a known read-only share is blocked.
 */
function canWritePersonalView(view: SaveTargetView, currentUserId: string | undefined, writeAccess: boolean | undefined): boolean {
	return !isSharedView(view, currentUserId) || writeAccess !== false;
}

/**
 * Whether the user can overwrite and publish the system view:
 *  - the view's IsCustomizable managed property allows updates (canOverwriteView, #93);
 *    https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#update-views
 *  - and the user can write and publish customizations (prvWriteQuery, prvWriteCustomization and
 *    prvPublishCustomization, as checkSavedQueryPrivileges reads them).
 * A managed but customizable view is offered: saving adds an unmanaged layer, which the confirmation
 * warns about (#142), as SaveViewDialog does today.
 */
function canCustomizeSystemView(view: SaveTargetView, privileges: SystemViewPrivileges | null | undefined): boolean {
	return canOverwriteView(view) && privileges?.canWrite === true && privileges.canPublish === true;
}

/**
 * The Save command for the loaded view (#278, Q13):
 *
 * | Loaded view                               | Primary                    | Menu                                         |
 * |-------------------------------------------|----------------------------|----------------------------------------------|
 * | None                                      | Save as new view…          | (none)                                       |
 * | Personal view with Write                  | Save changes to "{view}"   | Save as new view…                            |
 * | Personal view shared read-only            | Save as new view…          | (none)                                       |
 * | System view, user can customize it        | Save as new view…          | Save changes to system view "{view}"…        |
 * | System view, not customizable or no privs | Save as new view…          | (none)                                       |
 *
 * Overwriting a personal view someone else shared with Write asks first (it changes the view for
 * everyone it's shared with, #142). Overwriting a system view always asks, naming the view, then
 * publishes.
 */
export function saveOptions({ view, currentUserId, personalWriteAccess, systemViewPrivileges }: SaveOptionsInput): SaveOptions {
	if (!view) return { primary: SAVE_AS, menu: [] };

	if (view.type === "personal") {
		if (!canWritePersonalView(view, currentUserId, personalWriteAccess)) return { primary: SAVE_AS, menu: [] };
		return {
			primary: {
				kind: "overwrite-personal",
				label: `Save changes to ${quoted(view.name)}`,
				needsConfirm: isSharedView(view, currentUserId),
				publishes: false,
			},
			menu: [SAVE_AS],
		};
	}

	if (!canCustomizeSystemView(view, systemViewPrivileges)) return { primary: SAVE_AS, menu: [] };
	return {
		primary: SAVE_AS,
		menu: [
			{
				kind: "overwrite-system",
				label: `Save changes to system view ${quoted(view.name)}…`,
				needsConfirm: true,
				publishes: true,
			},
		],
	};
}
