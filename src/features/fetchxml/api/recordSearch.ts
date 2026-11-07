/**
 * Record search for the hierarchy operators' value picker (#98): the operators take the primary key
 * of a record in a hierarchical entity, so the picker finds records by their primary name.
 * https://learn.microsoft.com/power-apps/developer/data-platform/query-hierarchical-data
 *
 * - startswith, not contains: a leading wildcard can't use an index (wildcard-characters › Don't use
 *   leading wild cards), and Quick Find matches the start of the name by default too.
 * - The Web API's startswith treats % _ [ as wildcards (wildcard-characters lists it; verified live
 *   2026-10-17: contains(name,'%') matches every account), so typed text is escaped to match literally.
 * - The literal is URL-encoded: the host puts the query in the URL as given, so & # + would break it.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query/filter-rows#use-odata-query-functions
 */
import { escapeWildcards } from "../model/wildcards";
import { odataString } from "./odata";

/** Records listed by the picker */
export const RECORD_SEARCH_TOP = 10;

export interface RecordSearchTarget {
	entitySetName: string;
	idAttribute: string;
	nameAttribute: string;
}

/** Web API query for records whose primary name starts with the text (all records when it's empty). */
export function buildRecordSearchQuery(target: RecordSearchTarget, text: string, top = RECORD_SEARCH_TOP): string {
	const { entitySetName, idAttribute, nameAttribute } = target;
	const term = text.trim();
	const filter = term ? `&$filter=startswith(${nameAttribute},'${encodeURIComponent(odataString(escapeWildcards(term)))}')` : "";
	return `${entitySetName}?$select=${idAttribute},${nameAttribute}${filter}&$orderby=${nameAttribute}&$top=${top}`;
}

/** Web API query for one record's primary name. */
export function buildRecordNameQuery(target: RecordSearchTarget, id: string): string {
	return `${target.entitySetName}(${normalizeGuid(id)})?$select=${target.nameAttribute}`;
}

const HEX = "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}";
const GUID = new RegExp(`^(${HEX}|\\{${HEX}\\})$`, "i");

/**
 * Whether the value is a GUID, with or without braces. Dataverse rejects anything else for a
 * primary key or lookup value (verified live: "not-a-guid" with under fails with 0x80040203
 * System.FormatException; "{GUID}" in upper case runs).
 */
export const isGuid = (value: string): boolean => GUID.test(value.trim());

/** The GUID without braces or surrounding spaces. */
export const normalizeGuid = (value: string): string => value.trim().replace(/^\{|\}$/g, "");
