/**
 * Activate / Deactivate planning from the entity's state model (#143). Pure: the metadata comes in,
 * the target `statecode` / `statuscode` (or the choices for the user) come out.
 *
 * Microsoft Learn:
 * - A state change is an update of `statecode` and `statuscode` (the SetState message is deprecated;
 *   the Web API operation is `PATCH /{EntitySetName}(id)`), and you should always set the
 *   `statuscode` you want with the `statecode`:
 *   https://learn.microsoft.com/power-apps/developer/data-platform/special-update-operation-behavior
 *   https://learn.microsoft.com/power-apps/developer/data-platform/webapi/update-delete-entities-using-web-api#basic-update
 * - Each state option has a DefaultStatus; each status option has the State it belongs to:
 *   https://learn.microsoft.com/dotnet/api/microsoft.xrm.sdk.metadata.stateoptionmetadata.defaultstatus
 *   https://learn.microsoft.com/dotnet/api/microsoft.xrm.sdk.metadata.statusoptionmetadata.state
 * - With EntityMetadata.EnforceStateTransitions, each status option's TransitionData lists the
 *   statuses a record may move to from it ("Define custom state model transitions"):
 *   https://learn.microsoft.com/power-apps/developer/data-platform/define-custom-state-model-transitions
 *   https://learn.microsoft.com/power-apps/maker/data-platform/define-status-reason-transitions
 */

export interface StateOption {
	value: number;
	label: string;
	invariantName?: string;
	defaultStatus?: number;
}

export interface StatusOption {
	value: number;
	label: string;
	/** The statecode this status reason belongs to */
	state: number;
	/** Statuses allowed next (parsed TransitionData); undefined when the option has none */
	allowedNext?: number[];
}

export interface EntityStateModel {
	states: StateOption[];
	statuses: StatusOption[];
	/** EntityMetadata.EnforceStateTransitions */
	enforceTransitions: boolean;
}

export type StateAction = "activate" | "deactivate";

export interface StateTarget {
	statecode: number;
	statuscode: number;
}

/** The active state: Activate sets statecode 0, Deactivate one of the others (#143) */
export const ACTIVE_STATE = 0;

/**
 * Entities whose closing states go through dedicated messages, so Activate / Deactivate stay off
 * and name the documented action instead (#143). #149 adds dialogs for these messages.
 *
 * Verified on a sandbox (2026-10-17): PATCH of opportunity to Won / Lost, and of incident to
 * Resolved, is rejected with 0x80040216 ("This message can not be used to set the state of ... use
 * the ... message instead").
 */
export const CLOSE_MESSAGE_ENTITIES: Readonly<Record<string, { reason: string; docUrl: string }>> = {
	opportunity: {
		reason: "Opportunities close with Win or Lose (WinOpportunity / LoseOpportunity).",
		docUrl: "https://learn.microsoft.com/dynamics365/developer/reference/entities/opportunity#messages",
	},
	incident: {
		reason: "Cases are resolved with Resolve Case (CloseIncident).",
		docUrl: "https://learn.microsoft.com/dynamics365/developer/reference/entities/incident#messages",
	},
	quote: {
		reason: "Quotes close with Win or Close (WinQuote / CloseQuote).",
		docUrl: "https://learn.microsoft.com/dynamics365/developer/reference/entities/quote#messages",
	},
	salesorder: {
		reason: "Orders close with Fulfill or Cancel (FulfillSalesOrder / CancelSalesOrder).",
		docUrl: "https://learn.microsoft.com/dynamics365/developer/reference/entities/salesorder#messages",
	},
	invoice: {
		reason: "Invoices are closed as paid or canceled from the invoice itself.",
		docUrl: "https://learn.microsoft.com/dynamics365/developer/reference/entities/invoice#messages",
	},
	contract: {
		reason: "Contracts follow their own life cycle (Draft, Invoiced, Active), with CancelContract and RenewContract.",
		docUrl: "https://learn.microsoft.com/dynamics365/customerengagement/on-premises/developer/contract-entities",
	},
};

/** Why Activate / Deactivate are off for an entity whose states need dedicated messages, or undefined */
export function closeMessageReason(entityLogicalName: string | undefined): string | undefined {
	const entry = entityLogicalName ? CLOSE_MESSAGE_ENTITIES[entityLogicalName.toLowerCase()] : undefined;
	return entry ? `${entry.reason} Use the model-driven app for now.` : undefined;
}

/**
 * Targets that Dataverse accepts as a plain update but that skip what the app does for them. Shown
 * as a note when chosen. Verified on a sandbox (2026-10-17): PATCH of a lead to Qualified succeeds.
 */
export const STATE_NOTES: Readonly<Record<string, Readonly<Record<number, string>>>> = {
	lead: {
		1: "This sets the status only. It doesn't run Qualify (QualifyLead), so no account, contact or opportunity is created.",
	},
};

export const stateNote = (entityLogicalName: string | undefined, statecode: number): string | undefined =>
	entityLogicalName ? STATE_NOTES[entityLogicalName.toLowerCase()]?.[statecode] : undefined;

const XML_ENTITIES: Record<string, string> = { "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'", "&amp;": "&" };

/**
 * The `tostatusid` values of a TransitionData document. Learn notes the XML can arrive escaped, so
 * entities are decoded first. Undefined when there's no document.
 */
export function parseTransitionData(transitionData: string | null | undefined): number[] | undefined {
	if (!transitionData || !transitionData.trim()) return undefined;
	const xml = transitionData.replace(/&(lt|gt|quot|apos|amp);/g, (m) => XML_ENTITIES[m]);
	const targets: number[] = [];
	for (const match of xml.matchAll(/<allowedtransition\b[^>]*\btostatusid\s*=\s*["'](-?\d+)["']/gi)) {
		const value = Number(match[1]);
		if (!targets.includes(value)) targets.push(value);
	}
	return targets;
}

/** The statuses that belong to a state, in metadata order */
export const statusesOfState = (model: EntityStateModel, state: number) => model.statuses.filter((s) => s.state === state);

/** The states the action can move to: Activate the active state, Deactivate every other one */
export function targetStates(model: EntityStateModel, action: StateAction): StateOption[] {
	return model.states.filter((s) => (action === "activate" ? s.value === ACTIVE_STATE : s.value !== ACTIVE_STATE));
}

/**
 * Simple entities (#143): two states, one status reason each (account, contact). Activate and
 * Deactivate then need no choice.
 */
export function isSimpleStateModel(model: EntityStateModel): boolean {
	return model.states.length === 2 && model.states.every((s) => statusesOfState(model, s.value).length === 1);
}

/** The state's DefaultStatus, or its first status reason when metadata has no default */
function defaultStatusOf(model: EntityStateModel, state: StateOption): number | undefined {
	const statuses = statusesOfState(model, state.value);
	if (state.defaultStatus !== undefined && statuses.some((s) => s.value === state.defaultStatus)) return state.defaultStatus;
	return statuses[0]?.value;
}

export interface TransitionFilter {
	/** Status reasons allowed from every selected record's current status reason */
	allowed?: Set<number>;
	/** Shown above the choices: what the filter did, or why it couldn't run */
	note?: string;
}

/**
 * Status reason transitions (#143): with EnforceStateTransitions, only statuses allowed from the
 * current status of ALL selected records. Selected records with different current statuses get the
 * intersection. Statuses with no TransitionData, or records whose status couldn't be read, don't
 * restrict anything (block only what's known to be invalid; Dataverse rejects the rest).
 *
 * @param currentStatuses the selected records' current statuscode values; undefined if unknown
 */
export function transitionFilter(model: EntityStateModel, currentStatuses: number[] | undefined): TransitionFilter {
	if (!model.enforceTransitions) return {};
	if (!currentStatuses || currentStatuses.length === 0) {
		return { note: "Status reason transitions are enforced for this entity, but the records' current status reasons couldn't be read. Dataverse rejects targets that aren't allowed." };
	}
	const distinct = [...new Set(currentStatuses)];
	let allowed: Set<number> | undefined;
	for (const current of distinct) {
		const next = model.statuses.find((s) => s.value === current)?.allowedNext;
		if (!next) continue;
		const set = new Set(next);
		allowed = allowed ? new Set([...allowed].filter((v) => set.has(v))) : set;
	}
	if (!allowed) return {};
	return {
		allowed,
		note:
			distinct.length > 1
				? "Status reason transitions are enforced. The selected records have different status reasons, so only targets allowed from all of them are shown."
				: "Status reason transitions are enforced, so only targets allowed from the current status reason are shown.",
	};
}

export interface StatusChoice {
	statuscode: number;
	label: string;
	isDefault: boolean;
}

export interface StateGroup {
	statecode: number;
	label: string;
	choices: StatusChoice[];
}

/** The way back from a state change (#69): what the opposite action can set */
export interface StateUndo {
	/** The opposite action has a state to set */
	available: boolean;
	/** It offers more than one status reason, so the previous one isn't restored by itself */
	picksStatusReason: boolean;
	/** Status reason transitions are enforced, so going back may not be allowed */
	transitionsEnforced: boolean;
}

/** The way back from Activate or Deactivate: the opposite action's target states (#69) */
export function stateUndo(model: EntityStateModel, action: StateAction): StateUndo {
	const back = targetStates(model, action === "activate" ? "deactivate" : "activate");
	const reasons = back.flatMap((state) => statusesOfState(model, state.value));
	return { available: reasons.length > 0, picksStatusReason: reasons.length > 1, transitionsEnforced: model.enforceTransitions };
}

export type StatePlan =
	/** One target, no choice (simple entities) */
	| { kind: "simple"; target: StateTarget; stateLabel: string; statusLabel: string; undo: StateUndo }
	/** The user picks a status reason, grouped by state */
	| { kind: "choose"; groups: StateGroup[]; defaultTarget: StateTarget; note?: string; undo: StateUndo }
	/** Nothing to do, and why */
	| { kind: "none"; reason: string };

/**
 * What Activate or Deactivate sets for the selected records (#143). Simple entities get their one
 * target (the state's default status); others get the status reasons of the target states, grouped
 * by state. The preselected target is the default status of the first state without a STATE_NOTES
 * caveat (a lead's Deactivate preselects Disqualified, not Qualified).
 */
export function planStateChange(
	model: EntityStateModel,
	action: StateAction,
	currentStatuses?: number[],
	entityLogicalName?: string,
): StatePlan {
	const states = targetStates(model, action);
	if (states.length === 0) {
		return {
			kind: "none",
			reason: action === "activate" ? "This entity has no active state (statecode 0)." : "This entity has no inactive state.",
		};
	}

	if (isSimpleStateModel(model)) {
		const state = states[0];
		const statuscode = defaultStatusOf(model, state);
		if (statuscode === undefined) return { kind: "none", reason: `The ${state.label} state has no status reason.` };
		const statusLabel = model.statuses.find((s) => s.value === statuscode)?.label ?? String(statuscode);
		return {
			kind: "simple",
			target: { statecode: state.value, statuscode },
			stateLabel: state.label,
			statusLabel,
			undo: stateUndo(model, action),
		};
	}

	const filter = transitionFilter(model, currentStatuses);
	const groups: StateGroup[] = [];
	for (const state of states) {
		const def = defaultStatusOf(model, state);
		const choices = statusesOfState(model, state.value)
			.filter((s) => !filter.allowed || filter.allowed.has(s.value))
			.map((s) => ({ statuscode: s.value, label: s.label, isDefault: s.value === def }));
		if (choices.length > 0) groups.push({ statecode: state.value, label: state.label, choices });
	}
	if (groups.length === 0) {
		return {
			kind: "none",
			reason: filter.allowed
				? "Status reason transitions are enforced, and no target status reason is allowed from the selected records' current status reasons."
				: "The target states have no status reasons.",
		};
	}
	const first = groups.find((g) => !stateNote(entityLogicalName, g.statecode)) ?? groups[0];
	const preselected = first.choices.find((c) => c.isDefault) ?? first.choices[0];
	return {
		kind: "choose",
		groups,
		defaultTarget: { statecode: first.statecode, statuscode: preselected.statuscode },
		note: filter.note,
		undo: stateUndo(model, action),
	};
}

/** The body of the state change PATCH: `{ statecode, statuscode }` (Learn: always set both) */
export const statePatchBody = (target: StateTarget): Record<string, number> => ({
	statecode: target.statecode,
	statuscode: target.statuscode,
});

/** Metadata option shapes, as the Web API returns StateOptionMetadata / StatusOptionMetadata */
interface MetadataOption {
	Value: number;
	Label?: { UserLocalizedLabel?: { Label?: string } | null };
	InvariantName?: string;
	DefaultStatus?: number;
	State?: number;
	TransitionData?: string | null;
}

/** Builds the model from the Web API's StateAttributeMetadata and StatusAttributeMetadata options */
export function toEntityStateModel(
	stateOptions: MetadataOption[],
	statusOptions: MetadataOption[],
	enforceTransitions: boolean,
): EntityStateModel {
	const label = (o: MetadataOption) => o.Label?.UserLocalizedLabel?.Label || o.InvariantName || String(o.Value);
	return {
		states: stateOptions
			.map((o) => ({ value: o.Value, label: label(o), invariantName: o.InvariantName, defaultStatus: o.DefaultStatus ?? undefined }))
			.sort((a, b) => a.value - b.value),
		statuses: statusOptions
			.filter((o) => typeof o.State === "number")
			.map((o) => ({ value: o.Value, label: label(o), state: o.State as number, allowedNext: parseTransitionData(o.TransitionData) })),
		enforceTransitions,
	};
}

/**
 * Whether the action can apply at all, before records are chosen (the command bar): undefined when
 * it can, otherwise the reason it's off.
 */
export function stateActionUnavailableReason(model: EntityStateModel | null, action: StateAction): string | undefined {
	if (!model) return "This entity has no Status (statecode) column.";
	if (targetStates(model, action).length === 0) {
		return action === "activate" ? "This entity has no active state (statecode 0)." : "This entity has no inactive state.";
	}
	return undefined;
}
