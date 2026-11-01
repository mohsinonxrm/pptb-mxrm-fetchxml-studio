/**
 * Close actions for entities whose closing states go through dedicated messages (#149). Pure: the
 * action definitions per entity, the status reason choices from the state model, validation of the
 * dialog's values, and the request payload for one record.
 *
 * Microsoft Learn (SDK request classes; the Web API actions take the same parameter names):
 * - WinOpportunity / LoseOpportunity: OpportunityClose and Status required; Status -1 lets the system
 *   choose. https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.winopportunityrequest
 *   https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.loseopportunityrequest
 * - CloseIncident: IncidentResolution and Status required.
 *   https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.closeincidentrequest
 * - WinQuote / CloseQuote: QuoteClose and Status required; only active quotes close.
 *   https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.winquoterequest
 *   https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.closequoterequest
 * - FulfillSalesOrder / CancelSalesOrder: OrderClose and Status required. FulfillSalesOrder is an
 *   action bound to the salesorder collection ("Use Web API actions"):
 *   https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.fulfillsalesorderrequest
 *   https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.cancelsalesorderrequest
 *   https://learn.microsoft.com/power-apps/developer/data-platform/webapi/use-web-api-actions
 * - Entity parameters carry "@odata.type" so the action knows the close activity's type:
 *   https://learn.microsoft.com/power-apps/developer/data-platform/webapi/use-web-api-actions
 * - Close activity columns: https://learn.microsoft.com/dynamics365/developer/reference/entities/opportunityclose
 *   (actualrevenue Money precision 2, actualend DateOnly UserLocal, description 2,000, competitorid),
 *   .../incidentresolution (subject ApplicationRequired 200, timespent "Billable Time", description
 *   100,000), .../quoteclose, .../orderclose (description 2,000).
 * - Incident cancel and invoice paid / canceled are plain state changes (the entity references list
 *   SetState as PATCH of statecode and statuscode; invoices have no close message):
 *   https://learn.microsoft.com/dynamics365/developer/reference/entities/incident#messages
 *   https://learn.microsoft.com/dynamics365/developer/reference/entities/invoice#messages
 *
 * Verified on a sandbox with temporary records (2026-10-17): every message below with explicit and
 * -1 statuses; unbound FulfillSalesOrder answers 404 0x80060888, the collection-bound one 204; a
 * WinOpportunity without actualrevenue keeps the opportunity's own Actual revenue; incident cancel
 * (2/6) and invoice Paid (2/100001, 2/100002) and Canceled (3/100003) by PATCH return 200.
 */

import { stateUndo, transitionFilter, type EntityStateModel, type StateOption, type StatePlan, type StatusOption } from "./recordState";

export type CloseField = "revenue" | "date" | "description" | "competitor" | "resolution" | "billable";

interface CloseActionBase {
	/** Unique within the entity, e.g. "win" */
	key: string;
	/** Menu item and primary button, e.g. "Close as won" */
	label: string;
	/** Dialog title for the selected record count */
	title: (count: number) => string;
	/** The statecode the action moves records to */
	targetState: number;
	/** Shown above the fields */
	note?: string;
	/** Destructive outcome (lost, canceled): a divider goes above the first one in the menu */
	destructive?: boolean;
	/** Microsoft Learn page for the operation */
	docUrl: string;
}

/** A close message (WinOpportunity, CloseIncident, ...) sent once per record */
export interface CloseMessageAction extends CloseActionBase {
	kind: "message";
	/** Web API action name */
	message: string;
	/** Bound to the entity collection: POST {EntitySetName}/Microsoft.Dynamics.CRM.{message} */
	boundToCollection?: boolean;
	/** The close activity the message creates */
	closeEntity: string;
	/** The action parameter that takes the close activity */
	closeParameter: string;
	/** The close activity's lookup to the record (its navigation property, verified in metadata) */
	parentLookup: string;
	/** Status reason preselected instead of the state's DefaultStatus (quote: Lost, not Revised) */
	preselectStatus?: number;
	/** Status reason suggested once a competitor is named (opportunity Out-Sold), until the user picks one */
	competitorStatus?: number;
	/** Label for the status reason field */
	statusLabel?: string;
	fields: CloseField[];
	/** Maximum length of the close activity's description */
	descriptionMaxLength: number;
	/** Singular and plural nouns for the result, e.g. ["opportunity", "opportunities"] */
	nouns: [string, string];
}

/** A plain state change (PATCH statecode / statuscode) through the Activate / Deactivate dialog */
export interface CloseStateAction extends CloseActionBase {
	kind: "state";
	/** Status reasons of the target state not offered (incident Merged: Merge sets it) */
	excludeStatuses?: number[];
}

export type CloseAction = CloseMessageAction | CloseStateAction;

const SDK = "https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages";
const ENTITIES = "https://learn.microsoft.com/dynamics365/developer/reference/entities";

const plural = (count: number, one: string, many: (n: number) => string) => (count === 1 ? one : many(count));

/** Close actions per entity (#149), in menu order */
export const CLOSE_ACTIONS: Readonly<Record<string, readonly CloseAction[]>> = {
	opportunity: [
		{
			kind: "message",
			key: "win",
			label: "Close as won",
			title: (n) => plural(n, "Close opportunity as won", (c) => `Close ${c} opportunities as won`),
			message: "WinOpportunity",
			closeEntity: "opportunityclose",
			closeParameter: "OpportunityClose",
			parentLookup: "opportunityid",
			targetState: 1,
			fields: ["revenue", "date", "description"],
			descriptionMaxLength: 2000,
			nouns: ["opportunity", "opportunities"],
			docUrl: `${SDK}.winopportunityrequest`,
		},
		{
			kind: "message",
			key: "lose",
			label: "Close as lost",
			title: (n) => plural(n, "Close opportunity as lost", (c) => `Close ${c} opportunities as lost`),
			message: "LoseOpportunity",
			closeEntity: "opportunityclose",
			closeParameter: "OpportunityClose",
			parentLookup: "opportunityid",
			targetState: 2,
			// Out-Sold. LoseOpportunity with Status -1 and a competitor picks Canceled (verified), so the
			// dialog suggests it instead
			competitorStatus: 5,
			fields: ["competitor", "revenue", "date", "description"],
			descriptionMaxLength: 2000,
			nouns: ["opportunity", "opportunities"],
			destructive: true,
			docUrl: `${SDK}.loseopportunityrequest`,
		},
	],
	incident: [
		{
			kind: "message",
			key: "resolve",
			label: "Resolve case",
			title: (n) => plural(n, "Resolve case", (c) => `Resolve ${c} cases`),
			message: "CloseIncident",
			closeEntity: "incidentresolution",
			closeParameter: "IncidentResolution",
			parentLookup: "incidentid",
			targetState: 1,
			statusLabel: "Resolution type",
			fields: ["resolution", "billable", "description"],
			descriptionMaxLength: 100000,
			nouns: ["case", "cases"],
			docUrl: `${SDK}.closeincidentrequest`,
		},
		{
			kind: "state",
			key: "cancel",
			label: "Cancel case",
			title: (n) => plural(n, "Cancel case", (c) => `Cancel ${c} cases`),
			targetState: 2,
			excludeStatuses: [2000],
			note: "A plain state change, like Deactivate. Merged isn't offered: merging cases sets it.",
			destructive: true,
			docUrl: `${ENTITIES}/incident#messages`,
		},
	],
	quote: [
		{
			kind: "message",
			key: "win",
			label: "Win quote",
			title: (n) => plural(n, "Win quote", (c) => `Win ${c} quotes`),
			message: "WinQuote",
			closeEntity: "quoteclose",
			closeParameter: "QuoteClose",
			parentLookup: "quoteid",
			targetState: 2,
			fields: ["date", "description"],
			descriptionMaxLength: 2000,
			nouns: ["quote", "quotes"],
			note: "Only active quotes can be won or closed; a draft quote fails for that record. Winning doesn't create an order.",
			docUrl: `${SDK}.winquoterequest`,
		},
		{
			kind: "message",
			key: "close",
			label: "Close quote",
			title: (n) => plural(n, "Close quote", (c) => `Close ${c} quotes`),
			message: "CloseQuote",
			closeEntity: "quoteclose",
			closeParameter: "QuoteClose",
			parentLookup: "quoteid",
			targetState: 3,
			// Revised is the metadata default for Closed, but it normally comes from revising the quote
			preselectStatus: 5,
			fields: ["date", "description"],
			descriptionMaxLength: 2000,
			nouns: ["quote", "quotes"],
			destructive: true,
			note: "Only active quotes can be closed. Lost is preselected: Revised, the default, normally comes from revising the quote.",
			docUrl: `${SDK}.closequoterequest`,
		},
	],
	salesorder: [
		{
			kind: "message",
			key: "fulfill",
			label: "Fulfill order",
			title: (n) => plural(n, "Fulfill order", (c) => `Fulfill ${c} orders`),
			message: "FulfillSalesOrder",
			boundToCollection: true,
			closeEntity: "orderclose",
			closeParameter: "OrderClose",
			parentLookup: "salesorderid",
			targetState: 3,
			fields: ["date", "description"],
			descriptionMaxLength: 2000,
			nouns: ["order", "orders"],
			docUrl: `${SDK}.fulfillsalesorderrequest`,
		},
		{
			kind: "message",
			key: "cancel",
			label: "Cancel order",
			title: (n) => plural(n, "Cancel order", (c) => `Cancel ${c} orders`),
			message: "CancelSalesOrder",
			closeEntity: "orderclose",
			closeParameter: "OrderClose",
			parentLookup: "salesorderid",
			targetState: 2,
			fields: ["date", "description"],
			descriptionMaxLength: 2000,
			nouns: ["order", "orders"],
			destructive: true,
			docUrl: `${SDK}.cancelsalesorderrequest`,
		},
	],
	invoice: [
		{
			kind: "state",
			key: "paid",
			label: "Mark as paid",
			title: (n) => plural(n, "Mark invoice as paid", (c) => `Mark ${c} invoices as paid`),
			targetState: 2,
			note: "Invoices have no close message: this is a plain state change, like Deactivate.",
			docUrl: `${ENTITIES}/invoice#messages`,
		},
		{
			kind: "state",
			key: "cancel",
			label: "Cancel invoice",
			title: (n) => plural(n, "Cancel invoice", (c) => `Cancel ${c} invoices`),
			targetState: 3,
			note: "Invoices have no close message: this is a plain state change, like Deactivate.",
			destructive: true,
			docUrl: `${ENTITIES}/invoice#messages`,
		},
	],
};

/** The entity's close actions, or undefined when it has none (#149) */
export const closeActionsFor = (entityLogicalName: string | undefined): readonly CloseAction[] | undefined =>
	entityLogicalName ? CLOSE_ACTIONS[entityLogicalName.toLowerCase()] : undefined;

export const findCloseAction = (entityLogicalName: string | undefined, key: string): CloseAction | undefined =>
	closeActionsFor(entityLogicalName)?.find((a) => a.key === key);

export interface CloseStatusChoice {
	statuscode: number;
	label: string;
	isDefault: boolean;
}

export interface CloseStatusChoices {
	state?: StateOption;
	choices: CloseStatusChoice[];
	/** Preselected: the action's preselectStatus, else the state's DefaultStatus, else the first */
	preselected?: number;
}

/** The status reasons of the action's target state, from metadata (default marked) */
export function closeStatusChoices(model: EntityStateModel, action: CloseAction): CloseStatusChoices {
	const state = model.states.find((s) => s.value === action.targetState);
	const exclude = action.kind === "state" ? (action.excludeStatuses ?? []) : [];
	const statuses = model.statuses.filter((s: StatusOption) => s.state === action.targetState && !exclude.includes(s.value));
	const choices = statuses.map((s) => ({ statuscode: s.value, label: s.label, isDefault: s.value === state?.defaultStatus }));
	const has = (value: number | undefined) => value !== undefined && choices.some((c) => c.statuscode === value);
	const wanted = action.kind === "message" ? action.preselectStatus : undefined;
	const preselected = has(wanted) ? wanted : has(state?.defaultStatus) ? state?.defaultStatus : choices[0]?.statuscode;
	return { state, choices, preselected };
}

/**
 * The status reason to show while the user hasn't picked one: the action's competitorStatus once a
 * competitor is named (when metadata has it), otherwise the preselected one
 */
export function suggestedCloseStatus(action: CloseMessageAction, choices: CloseStatusChoices, competitorId: string | undefined): number | undefined {
	const competitorStatus = action.competitorStatus;
	if (competitorId && competitorStatus !== undefined && choices.choices.some((c) => c.statuscode === competitorStatus)) {
		return competitorStatus;
	}
	return choices.preselected;
}

/**
 * The Activate / Deactivate plan (#143) for a state close action (Cancel case, Mark as paid, Cancel
 * invoice): the target state's status reasons only, limited by enforced transitions like #143. One
 * choice needs none.
 */
export function planCloseStateChange(model: EntityStateModel, action: CloseStateAction, currentStatuses?: number[]): StatePlan {
	const { state, choices, preselected } = closeStatusChoices(model, action);
	if (!state || choices.length === 0) return { kind: "none", reason: "This entity has no status reason for this action." };
	const filter = transitionFilter(model, currentStatuses);
	const allowed = filter.allowed ? choices.filter((c) => filter.allowed!.has(c.statuscode)) : choices;
	if (allowed.length === 0) {
		return {
			kind: "none",
			reason: "Status reason transitions are enforced, and no target status reason is allowed from the selected records' current status reasons.",
		};
	}
	// The way back is Activate, as from Deactivate (#69)
	const undo = stateUndo(model, "deactivate");
	if (allowed.length === 1) {
		return {
			kind: "simple",
			target: { statecode: state.value, statuscode: allowed[0].statuscode },
			stateLabel: state.label,
			statusLabel: allowed[0].label,
			undo,
		};
	}
	const first = allowed.find((c) => c.statuscode === preselected) ?? allowed[0];
	return {
		kind: "choose",
		groups: [{ statecode: state.value, label: state.label, choices: allowed }],
		defaultTarget: { statecode: state.value, statuscode: first.statuscode },
		note: filter.note,
		undo,
	};
}

/** What the user enters in a close message dialog */
export interface CloseValues {
	statuscode: number;
	/** Actual revenue as typed; empty keeps each record's own value */
	revenue?: string;
	/** Close date (local calendar date) */
	date?: Date | null;
	description?: string;
	competitorId?: string;
	/** Case resolution: the incidentresolution subject */
	resolution?: string;
	/** Billable time in minutes */
	billableMinutes?: number | null;
}

/** Money's maximum (MoneyAttributeMetadata.MaxValue) */
const MONEY_MAX = 922337203685477;
/** IntegerAttributeMetadata max for timespent */
const INT_MAX = 2147483647;
export const RESOLUTION_MAX_LENGTH = 200;

/** Parses Actual revenue: empty is undefined, otherwise a non-negative amount with up to 2 decimals */
export function parseRevenue(text: string | undefined): number | undefined | "invalid" {
	const value = (text ?? "").trim();
	if (value === "") return undefined;
	if (!/^\d+(\.\d{1,2})?$/.test(value)) return "invalid";
	const amount = Number(value);
	return amount <= MONEY_MAX ? amount : "invalid";
}

/** Validation messages per field; empty when the values can be sent */
export function validateCloseValues(action: CloseMessageAction, values: CloseValues): Partial<Record<CloseField, string>> {
	const errors: Partial<Record<CloseField, string>> = {};
	if (action.fields.includes("revenue") && parseRevenue(values.revenue) === "invalid") {
		errors.revenue = "Enter an amount, like 25000 or 25000.50.";
	}
	if (action.fields.includes("date") && !values.date) errors.date = "Select a close date.";
	if (action.fields.includes("resolution") && !(values.resolution ?? "").trim()) errors.resolution = "Enter a resolution.";
	if (
		action.fields.includes("billable") &&
		values.billableMinutes !== null &&
		values.billableMinutes !== undefined &&
		(!Number.isInteger(values.billableMinutes) || values.billableMinutes < 0 || values.billableMinutes > INT_MAX)
	) {
		errors.billable = "Enter whole minutes, 0 or more.";
	}
	return errors;
}

/**
 * The close date for a DateOnly / UserLocal column (actualend on all four close activities): noon
 * of the chosen local day, in UTC. Dataverse shows it as a date in the user's time zone, so noon
 * keeps the same day when the browser's and the user's time zones differ by up to 11 hours.
 * Verified: "2026-10-17" (UTC midnight) shows as 10/16 for a user west of UTC.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/behavior-format-date-time-attribute
 */
export const closeDateValue = (date: Date): string =>
	new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12).toISOString();

/** The URL of the record, as an @odata.bind value */
const bindUrl = (entitySetName: string, recordId: string) => `/${entitySetName}(${recordId.replace(/[{}]/g, "")})`;

/** Entity set of competitor (system entity; verified in metadata) */
export const COMPETITOR_ENTITY_SET = "competitors";

/**
 * The action's parameters for one record: the close activity (typed with @odata.type, bound to
 * the record) and Status. Only fields the user filled in are sent; the same values go to every
 * record.
 */
export function buildCloseParameters(
	action: CloseMessageAction,
	entitySetName: string,
	recordId: string,
	values: CloseValues,
): Record<string, unknown> {
	const activity: Record<string, unknown> = {
		"@odata.type": `Microsoft.Dynamics.CRM.${action.closeEntity}`,
		[`${action.parentLookup}@odata.bind`]: bindUrl(entitySetName, recordId),
	};
	const fields = new Set(action.fields);
	if (fields.has("resolution")) activity.subject = (values.resolution ?? "").trim().slice(0, RESOLUTION_MAX_LENGTH);
	if (fields.has("revenue")) {
		const revenue = parseRevenue(values.revenue);
		if (typeof revenue === "number") activity.actualrevenue = revenue;
	}
	if (fields.has("date") && values.date) activity.actualend = closeDateValue(values.date);
	if (fields.has("billable") && typeof values.billableMinutes === "number") activity.timespent = values.billableMinutes;
	if (fields.has("competitor") && values.competitorId) {
		activity["competitorid@odata.bind"] = bindUrl(COMPETITOR_ENTITY_SET, values.competitorId);
	}
	const description = (values.description ?? "").trim();
	if (fields.has("description") && description) activity.description = description.slice(0, action.descriptionMaxLength);
	return { [action.closeParameter]: activity, Status: values.statuscode };
}

/** The request URL, relative to the Web API root: the action, or the collection-bound action */
export const closeRequestUrl = (action: CloseMessageAction, entitySetName: string): string =>
	action.boundToCollection ? `${entitySetName}/Microsoft.Dynamics.CRM.${action.message}` : action.message;

/**
 * The records a stopped $batch already closed: in the target state now and not before the run. A
 * record whose state before is unknown, or that already was in the target state, is redone and
 * Dataverse answers for it (block only what's known).
 */
export function completedByStoppedBatch(
	before: ReadonlyMap<string, number>,
	after: ReadonlyMap<string, number>,
	targetState: number,
	recordIds: string[],
): Set<string> {
	return new Set(
		recordIds.filter((id) => after.get(id) === targetState && before.has(id) && before.get(id) !== targetState),
	);
}

/** The result line: "1 opportunity closed." / "3 opportunities closed." */
export const closedSummary = (action: CloseMessageAction, count: number): string =>
	`${count} ${count === 1 ? action.nouns[0] : action.nouns[1]} closed.`;
