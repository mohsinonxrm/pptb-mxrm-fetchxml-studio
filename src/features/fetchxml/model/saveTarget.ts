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
 *
 * Checks the app can't complete don't hide a command: only a known "no" does. A check that failed
 * offers the command with a warning, and Dataverse decides when the user saves.
 */

import type { SavedView } from "../api/pptbClient";
import { canOverwriteView, isSharedView } from "./savedViews";

export type SaveActionKind = "overwrite-personal" | "overwrite-system" | "save-as";

export interface SaveAction {
	kind: SaveActionKind;
	/** The command's label, sentence case; "…" when it opens a dialog first */
	label: string;
	/** Ask before writing: a system view, a personal view someone else owns, or one whose owner isn't known */
	needsConfirm: boolean;
	/** Publish the Entity's customizations after the save (system views). A failed publish is the caller's soft error. */
	publishes: boolean;
	/** A check behind this command couldn't be completed: show this next to the command or in its confirmation */
	warning?: string;
}

export interface SaveOptions {
	primary: SaveAction;
	menu: SaveAction[];
}

/** The loaded view, as Load view reads it (pptbClient SavedView / builderStore LoadedViewState) */
export type SaveTargetView = Pick<SavedView, "type" | "name" | "isCustomizable" | "ownerId">;

/**
 * The result of one privilege or access check:
 *  - true / false: the check answered;
 *  - "unknown": the check failed (an error, or a response that can't be read);
 *  - undefined: the check hasn't answered yet.
 * The caller maps a failed check to "unknown". checkSavedQueryPrivileges and
 * checkUserQueryPrivileges report an error as false today, so wiring them up means returning
 * "unknown" from their catch blocks; checkUserQueryWriteAccess returns undefined for "couldn't
 * determine", which maps to "unknown" once its call has finished.
 */
export type CheckResult = boolean | "unknown" | undefined;

export interface SaveOptionsInput {
	/** The loaded view (the store's view identity); null when no view is loaded */
	view: SaveTargetView | null;
	/** The connected user (WhoAmI), to tell a personal view shared with the user from their own (#142); undefined while unknown */
	currentUserId?: string;
	/**
	 * Personal views: Write on the view itself (checkUserQueryWriteAccess: RetrievePrincipalAccess read
	 * with hasWriteAccess). Needed when someone else owns the view; false blocks the overwrite whoever owns it.
	 */
	personalWriteAccess?: CheckResult;
	/** Personal views: prvWriteUserQuery (checkUserQueryPrivileges().canWrite), which SaveViewDialog requires */
	canWriteUserQueries?: CheckResult;
	/**
	 * System views: prvWriteQuery and prvWriteCustomization (checkSavedQueryPrivileges().canWrite).
	 * prvPublishCustomization isn't a gate: Learn lists it as deprecated, and a failed publish ends as
	 * "Saved, but not published" with a retry (#180).
	 * https://learn.microsoft.com/power-platform/admin/miscellaneous-privileges
	 */
	canWriteSystemViews?: CheckResult;
}

/** A view name inside a label, quoted as in the owner decisions and today's Save tooltip */
const quoted = (name: string): string => `"${name}"`;

export const SAVE_AS_LABEL = "Save as new view…";
export const PRIVILEGES_UNKNOWN_WARNING =
	"FetchXML Studio couldn't check your privileges. Dataverse decides when you save.";
export const VIEW_ACCESS_UNKNOWN_WARNING =
	"FetchXML Studio couldn't check your access to this view. Dataverse decides when you save.";

const SAVE_AS: SaveAction = {
	kind: "save-as",
	label: SAVE_AS_LABEL,
	needsConfirm: false,
	publishes: false,
};
const SAVE_AS_ONLY: SaveOptions = { primary: SAVE_AS, menu: [] };

/** Whether a check blocks the command: a known "no", or no answer yet (so the command doesn't flash in and out) */
const blocks = (check: CheckResult): boolean => check === false || check === undefined;

/** Ownership of a personal view: the user's own, someone else's (#142), or unknown (an ID missing) */
function ownership(
	view: SaveTargetView,
	currentUserId: string | undefined,
): "own" | "shared" | "unknown" {
	if (!view.ownerId || !currentUserId) return "unknown";
	return isSharedView(view, currentUserId) ? "shared" : "own";
}

/**
 * Personal view: overwrite needs prvWriteUserQuery and Write on the view.
 *  - Write known false: Save as only, whoever owns the view (a read-only share, #142).
 *  - The user's own view: Write is theirs; overwrite without a confirmation.
 *  - Someone else's: overwrite once Write is true or "unknown", and confirm (it changes the view
 *    for everyone it's shared with).
 *  - Owner unknown (WhoAmI or ownerid missing): overwrite, and confirm.
 */
function personalOptions(input: SaveOptionsInput, view: SaveTargetView): SaveOptions {
	const { currentUserId, personalWriteAccess, canWriteUserQueries } = input;
	if (personalWriteAccess === false || blocks(canWriteUserQueries)) return SAVE_AS_ONLY;
	const owner = ownership(view, currentUserId);
	if (owner === "shared" && personalWriteAccess === undefined) return SAVE_AS_ONLY;

	const warning =
		canWriteUserQueries === "unknown"
			? PRIVILEGES_UNKNOWN_WARNING
			: owner !== "own" && personalWriteAccess === "unknown"
				? VIEW_ACCESS_UNKNOWN_WARNING
				: undefined;
	return {
		primary: {
			kind: "overwrite-personal",
			label: `Save changes to ${quoted(view.name)}`,
			needsConfirm: owner !== "own",
			publishes: false,
			...(warning ? { warning } : {}),
		},
		menu: [SAVE_AS],
	};
}

/**
 * System view: the menu offers overwrite-and-publish when
 *  - the view's IsCustomizable managed property allows updates (canOverwriteView, #93; unknown is allowed);
 *    https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#update-views
 *  - and the user can write system views (true, or "unknown" with a warning).
 * A managed but customizable view is offered: saving adds an unmanaged layer, which the confirmation
 * warns about (#142), as SaveViewDialog does today.
 */
function systemOptions(input: SaveOptionsInput, view: SaveTargetView): SaveOptions {
	const { canWriteSystemViews } = input;
	if (!canOverwriteView(view) || blocks(canWriteSystemViews)) return SAVE_AS_ONLY;
	return {
		primary: SAVE_AS,
		menu: [
			{
				kind: "overwrite-system",
				label: `Save changes to system view ${quoted(view.name)}…`,
				needsConfirm: true,
				publishes: true,
				...(canWriteSystemViews === "unknown"
					? { warning: PRIVILEGES_UNKNOWN_WARNING }
					: {}),
			},
		],
	};
}

/**
 * The Save command for the loaded view (#278, Q13):
 *
 * | Loaded view                                       | Primary                  | Menu                                  |
 * |---------------------------------------------------|--------------------------|---------------------------------------|
 * | None                                              | Save as new view…        | (none)                                |
 * | Personal view with Write and prvWriteUserQuery    | Save changes to "{view}" | Save as new view…                     |
 * | Personal view shared read-only, or no privilege   | Save as new view…        | (none)                                |
 * | System view, customizable, user can write it      | Save as new view…        | Save changes to system view "{view}"… |
 * | System view, not customizable, or can't write     | Save as new view…        | (none)                                |
 *
 * A check still loading hides the overwrite; a check that failed ("unknown") offers it with a warning.
 */
export function saveOptions(input: SaveOptionsInput): SaveOptions {
	const { view } = input;
	if (!view) return SAVE_AS_ONLY;
	return view.type === "personal" ? personalOptions(input, view) : systemOptions(input, view);
}
