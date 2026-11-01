/**
 * PPTB Dataverse API client wrapper
 * Provides typed interface to window.dataverseAPI exposed by PPTB host
 */

// NOTE: window.dataverseAPI global types are provided by @pptb/types (see tsconfig.app.json)

// Extended type for accessing Dataverse OData response annotations beyond what @pptb/types explicitly declares.
// Dataverse returns these extra annotations on FetchXML responses; we cast to this locally.
interface DataverseFetchXmlResponse {
	value: Record<string, unknown>[];
	"@odata.context"?: string;
	"@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"?: string;
	"@Microsoft.Dynamics.CRM.totalrecordcount"?: number;
	"@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"?: boolean;
	"@Microsoft.Dynamics.CRM.morerecords"?: boolean;
	/** Not sent by the Web API (verified); kept only as a defensive fallback */
	"@Microsoft.Dynamics.CRM.pagingcookie"?: string;
	/** Saved views (?savedQuery= / ?userQuery=) page with nextLink instead */
	"@odata.nextLink"?: string;
}

import { debugLog } from "../../../shared/utils/debug";
import { responseHasFormattedValues } from "./formattedValues";
import { metadataCache } from "../state/cache";
import { nextLinkToQueryPath, parseFetchXmlPagingCookie } from "./paging";
import { formatDataverseError } from "./dataverseErrors";
import { withThrottleRetry } from "./throttleRetry";
import { memoizeForSession } from "./sessionCache";
import { selectedRecordsQuery, withStringConditionValues } from "./bulkDeleteQuery";

// Connection-change hook (#71): forget the memoized WhoAmI and privilege answers
export { resetDataverseSessionCache } from "./sessionCache";
import { hasWriteAccess, readManagedPropertyValue, sortViewsByQueryType } from "../model/savedViews";
import { canExecuteBatch, connectionGuard, getPrimaryConnection } from "./host";
import {
	closeRecordsInBatches,
	deleteRecordsInBatches,
	updateRecordStatesInBatches,
	CLOSE_CONNECTION_CHANGED_MESSAGE,
	CONNECTION_CHANGED_MESSAGE,
	STATE_CONNECTION_CHANGED_MESSAGE,
	type ExecuteBatch,
} from "./batchDelete";
import { statePatchBody, toEntityStateModel, type EntityStateModel, type StateTarget } from "../model/recordState";
import {
	buildCloseParameters,
	closeRequestUrl,
	completedByStoppedBatch,
	COMPETITOR_ENTITY_SET,
	type CloseMessageAction,
	type CloseValues,
} from "../model/closeActions";

import { functionWithString, odataString } from "./odata";
import { buildRecordNameQuery, buildRecordSearchQuery, type RecordSearchTarget } from "./recordSearch";
import { USER_SETTINGS_TIME_ZONE_COLUMNS, timeZoneFromUserSettings, type DataverseTimeZone } from "../model/dateTimeValues";
export interface EntityMetadata {
	LogicalName: string;
	SchemaName: string;
	DisplayName?: { UserLocalizedLabel?: { Label?: string } };
	EntitySetName: string;
	PrimaryIdAttribute: string;
	PrimaryNameAttribute: string;
	IsValidForAdvancedFind?: boolean;
	MetadataId: string;
	ObjectTypeCode: number;
	/** Lower-case collection name; routes when a few system entities' EntitySetName doesn't (#61) */
	LogicalCollectionName?: string;
	/** Long-term retention is enabled: datasource="retained" can query it (#90b) */
	IsRetentionEnabled?: boolean;
}

export interface AttributeMetadata {
	LogicalName: string;
	SchemaName: string;
	DisplayName?: { UserLocalizedLabel?: { Label?: string } };
	AttributeType: string;
	AttributeTypeName?: { Value?: string };
	MetadataId: string;
	IsValidForAdvancedFind?: { Value: boolean };
	/** False: the value can't be retrieved, so any query using the attribute fails (#57) */
	IsValidForRead?: boolean;
	/** Logical attributes include composite ones (address1_composite), which can't be grouped or aggregated */
	IsLogical?: boolean;
	IsPrimaryId?: boolean;
	/**
	 * 0 simple, 1 calculated, 2 rollup, 3 formula, 4 prompt; null for types that can't be one (#97)
	 * https://learn.microsoft.com/power-apps/developer/data-platform/specialized-columns#common-elements-and-characteristics
	 */
	SourceType?: number | null;
	MaxLength?: number;
	Precision?: number;
	// For Integer, BigInt, Decimal, Double attributes
	MinValue?: number;
	MaxValue?: number;
	// For DateTime attributes
	Format?: "DateOnly" | "DateAndTime";
	DateTimeBehavior?: { Value?: string }; // Can be "UserLocal", "DateOnly", "TimeZoneIndependent"
	MinSupportedValue?: string; // ISO date string
	MaxSupportedValue?: string; // ISO date string
	// For Picklist, State, Status attributes
	OptionSet?: {
		Options?: Array<{
			Value: number;
			Label: { UserLocalizedLabel?: { Label?: string } };
		}>;
		// For Boolean/TwoOptions attributes
		TrueOption?: {
			Value: number;
			Label: { UserLocalizedLabel?: { Label?: string } };
		};
		FalseOption?: {
			Value: number;
			Label: { UserLocalizedLabel?: { Label?: string } };
		};
	};
	Targets?: string[]; // For Lookup attributes
}

export interface RelationshipMetadata {
	SchemaName: string;
	RelationshipType: "OneToManyRelationship" | "ManyToOneRelationship" | "ManyToManyRelationship";
	ReferencedEntity: string;
	ReferencedAttribute: string;
	ReferencingEntity: string;
	ReferencingAttribute: string;
	MetadataId: string;
	IsValidForAdvancedFind?: boolean;
	IsCustomRelationship?: boolean;
	// N-N specific properties
	Entity1LogicalName?: string;
	Entity2LogicalName?: string;
	Entity1IntersectAttribute?: string;
	Entity2IntersectAttribute?: string;
	IntersectEntityName?: string;
	/**
	 * Set by FetchXML Studio on the many-to-one joins it derives for an intersect entity: which side of
	 * its many-to-many relationship the join goes to (#57). Not Dataverse metadata.
	 */
	IntersectSide?: 1 | 2;
}

export interface FetchXmlResult {
	records: Record<string, unknown>[];
	totalRecordCount?: number;
	totalRecordCountLimitExceeded?: boolean;
	moreRecords?: boolean;
	/** Decoded paging cookie for the next page; undefined when Dataverse returned none (simple paging) */
	pagingCookie?: string;
	/** Saved views: the link to the next page */
	nextLink?: string;
}

/** Paging fields of a Web API response (#58): the documented annotations, cookie decoded. */
function pagingFromResponse(result: DataverseFetchXmlResponse): Pick<FetchXmlResult, "moreRecords" | "pagingCookie" | "nextLink"> {
	const fromAnnotation = parseFetchXmlPagingCookie(result["@Microsoft.Dynamics.CRM.fetchxmlpagingcookie"]);
	const legacy = result["@Microsoft.Dynamics.CRM.pagingcookie"];
	const nextLink = result["@odata.nextLink"];
	return {
		// An empty page ends paging even if more is claimed, so a loop can never spin on it
		moreRecords: (result["@Microsoft.Dynamics.CRM.morerecords"] === true || !!nextLink) && (result.value?.length ?? 0) > 0,
		pagingCookie: fromAnnotation.cookie ?? (typeof legacy === "string" && legacy ? legacy : undefined),
		nextLink: typeof nextLink === "string" ? nextLink : undefined,
	};
}

export interface WhoAmIResponse {
	UserId: string;
	BusinessUnitId: string;
	OrganizationId: string;
}

export interface PrivilegeCheckResponse {
	RolePrivileges: Array<{
		Depth: "Basic" | "Local" | "Deep" | "Global";
		PrivilegeId: string;
		BusinessUnitId: string;
		PrivilegeName: string;
	}>;
}

export interface Publisher {
	publisherid: string;
	friendlyname: string;
	uniquename: string;
	customizationprefix: string;
}

export interface Solution {
	solutionid: string;
	friendlyname: string;
	uniquename: string;
	solutionpackageversion?: string;
	version?: string;
	isvisible?: boolean;
	ismanaged?: boolean;
	_publisherid_value?: string;
}

export interface PublisherWithSolutions {
	publisher: Publisher;
	solutions: Solution[];
}

export interface SolutionComponent {
	msdyn_name: string; // Entity logical name
	msdyn_displayname?: string;
	msdyn_logicalcollectionname?: string;
	msdyn_solutionid: string;
	msdyn_componenttype: number;
}

export interface AccessSummary {
	userId: string;
	canReadPublisher: boolean;
	canReadSolution: boolean;
	canReadCustomization: boolean;
	fullFilterMode: boolean; // all 3 privileges
	solutionsOnlyMode: boolean; // customization + solution, no publisher
	publishersOnlyMode: boolean; // customization + publisher, no solution
	metadataOnlyMode: boolean; // customization only
	noAccessMode: boolean; // no customization privilege
}

/**
 * Check if PPTB Dataverse API is available
 */
export function isDataverseAvailable(): boolean {
	return typeof window !== "undefined" && !!window.dataverseAPI;
}

/**
 * Get all entities metadata (filtered for Advanced Find if specified)
 */
export async function getAllEntities(advancedFindOnly: boolean = true): Promise<EntityMetadata[]> {
	if (!isDataverseAvailable()) {
		const error = new Error("PPTB Dataverse API not available");
		console.error("pptbClient.getAllEntities:", error.message);
		throw error;
	}

	try {
		// Build OData query with server-side filtering
		let query =
			"EntityDefinitions?$select=LogicalName,SchemaName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,IsValidForAdvancedFind,MetadataId,ObjectTypeCode";

		// Add server-side filter for IsValidForAdvancedFind
		if (advancedFindOnly) {
			query += "&$filter=IsValidForAdvancedFind eq true";
		}

		debugLog("metadataAPI", `📡 GET Entities: ${query}`);

		const result = await window.dataverseAPI!.queryData(query);

		const entities = result.value as unknown as EntityMetadata[];

		debugLog("metadataAPI", `✅ Entities retrieved: ${entities.length} entities`);

		// Sort alphabetically by display name (fallback to logical name)
		// Must be client-side - $orderby not supported for metadata queries
		entities.sort((a, b) => {
			const nameA = a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
			const nameB = b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
			return nameA.localeCompare(nameB);
		});

		return entities;
	} catch (error) {
		console.error("pptbClient.getAllEntities: API call failed:", error);
		throw error;
	}
}

/**
 * Get entity metadata by logical name
 */
export async function getEntityMetadata(logicalName: string): Promise<EntityMetadata> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	return (await window.dataverseAPI!.getEntityMetadata(
		logicalName,
		true,
	)) as unknown as EntityMetadata;
}

/**
 * Get entity attributes metadata (filtered server-side for Advanced Find and sorted alphabetically)
 */
export async function getEntityAttributes(
	logicalName: string,
	advancedFindOnly: boolean = true,
): Promise<AttributeMetadata[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// Build OData query with server-side filtering
	let query = `EntityDefinitions(LogicalName='${odataString(logicalName)}')/Attributes?$select=LogicalName,SchemaName,DisplayName,AttributeType,AttributeTypeName,MetadataId,IsValidForAdvancedFind,IsValidForRead,IsLogical,IsPrimaryId,SourceType`;

	// Add server-side filter for IsValidForAdvancedFind
	if (advancedFindOnly) {
		query += "&$filter=IsValidForAdvancedFind/Value eq true";
	}

	debugLog("metadataAPI", `📡 GET Attributes for '${logicalName}': ${query}`);

	const result = await window.dataverseAPI!.queryData(query);

	const attributes = result.value as unknown as AttributeMetadata[];

	debugLog(
		"metadataAPI",
		`✅ Attributes retrieved for '${logicalName}': ${attributes.length} attributes`,
	);

	// Sort alphabetically by display name (fallback to logical name)
	// Must be client-side - $orderby not supported for metadata queries
	attributes.sort((a, b) => {
		const nameA = a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
		const nameB = b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
		return nameA.localeCompare(nameB);
	});

	return attributes;
}

/**
 * MaxLength of every String attribute of an entity (#97: conditions on String attributes longer
 * than 850 characters are slow). MaxLength is only on the String / Memo types, so it's read with
 * the StringAttributeMetadata cast; only LogicalName and MaxLength are returned.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query-metadata-web-api
 */
export async function getStringAttributeMaxLengths(logicalName: string): Promise<AttributeMetadata[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	const query = `EntityDefinitions(LogicalName='${odataString(logicalName)}')/Attributes/Microsoft.Dynamics.CRM.StringAttributeMetadata?$select=LogicalName,MaxLength`;
	debugLog("metadataAPI", `📡 GET String MaxLength for '${logicalName}': ${query}`);
	const result = await window.dataverseAPI!.queryData(query);
	return result.value as unknown as AttributeMetadata[];
}

/**
 * Get a single attribute with full metadata including OptionSet expansion
 * Required for loading picklist/boolean option values
 * Uses type casting (BooleanAttributeMetadata or PicklistAttributeMetadata) to access OptionSet
 */
export async function getAttributeWithOptionSet(
	entityLogicalName: string,
	attributeLogicalName: string,
): Promise<AttributeMetadata> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// First, get the basic attribute to determine its type
	const basicQuery = `EntityDefinitions(LogicalName='${odataString(entityLogicalName)}')/Attributes(LogicalName='${odataString(attributeLogicalName)}')?$select=AttributeType,AttributeTypeName,LogicalName,SchemaName,DisplayName,MetadataId`;

	debugLog("metadataAPI", `📡 GET Attribute type for '${attributeLogicalName}': ${basicQuery}`);

	const basicResult = await window.dataverseAPI!.queryData(basicQuery);
	// Single attribute query returns object directly, not in .value array
	const basicAttr = (basicResult.value?.[0] || basicResult) as unknown as AttributeMetadata;

	if (!basicAttr || !basicAttr.LogicalName) {
		throw new Error(`Attribute ${attributeLogicalName} not found in entity ${entityLogicalName}`);
	}

	// Determine the type cast needed based on AttributeType
	let typeCast = "";
	if (basicAttr.AttributeType === "Boolean") {
		typeCast = "Microsoft.Dynamics.CRM.BooleanAttributeMetadata";
	} else if (basicAttr.AttributeType === "State") {
		typeCast = "Microsoft.Dynamics.CRM.StateAttributeMetadata";
	} else if (basicAttr.AttributeType === "Status") {
		typeCast = "Microsoft.Dynamics.CRM.StatusAttributeMetadata";
	} else if (basicAttr.AttributeType === "Picklist") {
		typeCast = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata";
	} else if (basicAttr.AttributeTypeName?.Value === "MultiSelectPicklistType") {
		// Multi-select choices: contain-values / not-contain-values take their option values (#87)
		typeCast = "Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata";
	} else {
		// For non-optionset attributes, return the basic metadata
		debugLog(
			"metadataAPI",
			`⚠️ Attribute '${attributeLogicalName}' is type '${basicAttr.AttributeType}' - no OptionSet available`,
		);
		return basicAttr;
	}

	// Build query with type cast and OptionSet expansion
	const fullQuery = `EntityDefinitions(LogicalName='${odataString(entityLogicalName)}')/Attributes(LogicalName='${odataString(attributeLogicalName)}')/${typeCast}?$expand=OptionSet`;

	debugLog(
		"metadataAPI",
		`📡 GET Attribute with OptionSet for '${attributeLogicalName}': ${fullQuery}`,
	);

	const fullResult = await window.dataverseAPI!.queryData(fullQuery);
	// Single attribute query returns object directly, not in .value array
	const fullAttr = (fullResult.value?.[0] || fullResult) as unknown as AttributeMetadata;

	if (!fullAttr || !fullAttr.LogicalName) {
		throw new Error(
			`Failed to retrieve attribute ${attributeLogicalName} with OptionSet expansion`,
		);
	}

	debugLog(
		"metadataAPI",
		`✅ Attribute with OptionSet retrieved for '${attributeLogicalName}':`,
		fullAttr.OptionSet,
	);

	return fullAttr;
}

/**
 * Get detailed attribute metadata with type-specific properties (MinValue, MaxValue, Precision, Format)
 * Used for attributes that need validation constraints: Integer, BigInt, Decimal, Double, DateTime
 * Uses type casting (e.g., DecimalAttributeMetadata) to retrieve type-specific properties
 */
export async function getAttributeDetailedMetadata(
	entityLogicalName: string,
	attributeLogicalName: string,
	attributeType: string,
): Promise<AttributeMetadata> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// Determine the type cast and $select based on AttributeType
	// NOTE: When using type casting, only select type-specific properties (not base AttributeMetadata properties)
	// This is a Dataverse API limitation - see: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-metadata-web-api
	let typeCast = "";
	let selectProps = "";

	switch (attributeType) {
		case "Integer":
			typeCast = "Microsoft.Dynamics.CRM.IntegerAttributeMetadata";
			selectProps = "SchemaName,MaxValue,MinValue,Format";
			break;
		case "BigInt":
			typeCast = "Microsoft.Dynamics.CRM.BigIntAttributeMetadata";
			selectProps = "SchemaName,MaxValue,MinValue";
			break;
		case "Decimal":
			typeCast = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata";
			selectProps = "SchemaName,MaxValue,MinValue,Precision";
			break;
		case "Double":
		case "Float": // Float is an alias for Double in Dataverse
			typeCast = "Microsoft.Dynamics.CRM.DoubleAttributeMetadata";
			selectProps = "SchemaName,MaxValue,MinValue,Precision";
			break;
		case "DateTime":
			typeCast = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata";
			selectProps = "SchemaName,Format,DateTimeBehavior";
			break;
		default: {
			// For types that don't need detailed metadata, return early with warning
			debugLog(
				"metadataAPI",
				`⚠️ Attribute '${attributeLogicalName}' is type '${attributeType}' - no detailed metadata needed`,
			);
			// Return basic metadata query without type cast
			const basicQuery = `EntityDefinitions(LogicalName='${odataString(entityLogicalName)}')/Attributes(LogicalName='${odataString(attributeLogicalName)}')?$select=SchemaName,LogicalName,DisplayName,AttributeType,MetadataId`;
			const basicResult = await window.dataverseAPI!.queryData(basicQuery);
			return (basicResult.value?.[0] || basicResult) as unknown as AttributeMetadata;
		}
	}

	// Build query with type cast and specific property selection
	const detailedQuery = `EntityDefinitions(LogicalName='${odataString(entityLogicalName)}')/Attributes(LogicalName='${odataString(attributeLogicalName)}')/${typeCast}?$select=${selectProps}`;

	debugLog(
		"metadataAPI",
		`📡 GET Detailed Attribute metadata for '${attributeLogicalName}' (${attributeType}): ${detailedQuery}`,
	);

	const result = await window.dataverseAPI!.queryData(detailedQuery);
	// Single attribute query returns object directly, not in .value array
	const detailedAttr = (result.value?.[0] || result) as unknown as AttributeMetadata;

	if (!detailedAttr || !detailedAttr.SchemaName) {
		throw new Error(`Failed to retrieve detailed metadata for attribute ${attributeLogicalName}`);
	}

	// Add LogicalName back since it's not included in $select with type casting
	detailedAttr.LogicalName = attributeLogicalName;
	detailedAttr.AttributeType = attributeType;

	debugLog(
		"metadataAPI",
		`✅ Detailed attribute metadata retrieved for '${attributeLogicalName}':`,
		detailedAttr,
	);

	return detailedAttr;
}

/**
 * Get entity relationships metadata (filtered server-side for Advanced Find and sorted alphabetically)
 */
export async function getEntityRelationships(
	logicalName: string,
	advancedFindOnly: boolean = true,
): Promise<{
	oneToMany: RelationshipMetadata[];
	manyToOne: RelationshipMetadata[];
	manyToMany: RelationshipMetadata[];
}> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// Build queries with server-side filtering
	const buildQuery = (relType: string) => {
		let query = `EntityDefinitions(LogicalName='${odataString(logicalName)}')/${relType}?$select=SchemaName,RelationshipType,ReferencedEntity,ReferencedAttribute,ReferencingEntity,ReferencingAttribute,MetadataId,IsValidForAdvancedFind,IsCustomRelationship`;

		// Add server-side filter for IsValidForAdvancedFind
		if (advancedFindOnly) {
			query += "&$filter=IsValidForAdvancedFind eq true";
		}

		return query;
	};

	const buildManyToManyQuery = () => {
		let query = `EntityDefinitions(LogicalName='${odataString(logicalName)}')/ManyToManyRelationships?$select=SchemaName,RelationshipType,Entity1LogicalName,Entity2LogicalName,Entity1IntersectAttribute,Entity2IntersectAttribute,IntersectEntityName,MetadataId,IsValidForAdvancedFind,IsCustomRelationship`;

		// Add server-side filter for IsValidForAdvancedFind
		if (advancedFindOnly) {
			query += "&$filter=IsValidForAdvancedFind eq true";
		}

		return query;
	};

	const oneToManyQuery = buildQuery("OneToManyRelationships");
	const manyToOneQuery = buildQuery("ManyToOneRelationships");
	const manyToManyQuery = buildManyToManyQuery();

	debugLog("metadataAPI", `📡 GET 1:N Relationships for '${logicalName}': ${oneToManyQuery}`);
	debugLog("metadataAPI", `📡 GET N:1 Relationships for '${logicalName}': ${manyToOneQuery}`);
	debugLog("metadataAPI", `📡 GET N:N Relationships for '${logicalName}': ${manyToManyQuery}`);

	const [oneToManyResult, manyToOneResult, manyToManyResult] = await Promise.all([
		window.dataverseAPI!.queryData(oneToManyQuery),
		window.dataverseAPI!.queryData(manyToOneQuery),
		window.dataverseAPI!.queryData(manyToManyQuery),
	]);

	debugLog(
		"metadataAPI",
		`✅ Relationships retrieved for '${logicalName}': 1:N=${oneToManyResult.value.length}, N:1=${manyToOneResult.value.length}, N:N=${manyToManyResult.value.length}`,
	);

	// Helper function to sort relationships alphabetically by SchemaName
	// Must be client-side - $orderby not supported for metadata queries
	const sortRelationships = (rels: unknown[]): RelationshipMetadata[] => {
		const relationships = rels as RelationshipMetadata[];
		relationships.sort((a, b) => a.SchemaName.localeCompare(b.SchemaName));
		return relationships;
	};

	return {
		oneToMany: sortRelationships(oneToManyResult.value),
		manyToOne: sortRelationships(manyToOneResult.value),
		manyToMany: sortRelationships(manyToManyResult.value),
	};
}

/**
 * Root <entity> logical name of a FetchXML string, or undefined when it can't be determined
 * (malformed XML is left for Dataverse to report).
 */
export function getFetchXmlEntityName(fetchXml: string): string | undefined {
	const doc = new DOMParser().parseFromString(fetchXml, "text/xml");
	if (doc.getElementsByTagName("parsererror").length > 0) return undefined;
	const root = doc.documentElement;
	if (root?.nodeName !== "fetch") return undefined;
	for (const child of Array.from(root.children)) {
		if (child.nodeName === "entity") return child.getAttribute("name") || undefined;
	}
	return undefined;
}

/**
 * EntitySetName from metadata (cached) for the FetchXML's root entity (#61). Microsoft: the entity
 * set name must be the one for the table in the FetchXML entity element. The host's fetchXmlQuery
 * guesses it with an English pluralizer, which fails for entities such as `feedback` (verified live).
 */
async function resolveFetchXmlEntitySet(
	fetchXml: string,
): Promise<{ entitySetName: string; logicalCollectionName?: string } | undefined> {
	const logicalName = getFetchXmlEntityName(fetchXml);
	if (!logicalName) return undefined;
	return resolveEntitySet(logicalName);
}

/**
 * EntitySetName (and LogicalCollectionName) of a table from metadata (cached), for Web API URLs
 * the host would otherwise build with its pluralizer (#61). Undefined when metadata has none.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/web-api-service-documents#entity-set-name
 */
export async function resolveEntitySet(
	logicalName: string,
): Promise<{ entitySetName: string; logicalCollectionName?: string } | undefined> {
	let metadata = metadataCache.getEntityMetadata(logicalName);
	if (!metadata) {
		const isCurrent = metadataCache.generationGuard();
		metadata = await getEntityMetadata(logicalName);
		if (isCurrent()) metadataCache.setEntityMetadata(logicalName, metadata);
	}
	if (!metadata.EntitySetName) return undefined;
	return { entitySetName: metadata.EntitySetName, logicalCollectionName: metadata.LogicalCollectionName };
}

/** Dataverse's answer when the URL's entity set segment doesn't exist (0x80060888). */
const isUnknownSegmentError = (error: unknown) =>
	error instanceof Error && /Resource not found for the segment/i.test(error.message);

/**
 * Whether the entity has a hierarchical relationship (a self-referencing one-to-many relationship
 * with IsHierarchical = true). Count child records (rowaggregate) only works on these entities
 * (verified: 7 of 7 hierarchical entities accept it, 5 of 5 others reject it with 0x80047020).
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-hierarchical-data
 */
export async function getIsHierarchical(logicalName: string): Promise<boolean> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	const result = await window.dataverseAPI!.queryData(
		`EntityDefinitions(LogicalName='${odataString(logicalName)}')/OneToManyRelationships?$select=SchemaName&$filter=IsHierarchical eq true`,
	);
	return result.value.length > 0;
}

/**
 * The entity's hierarchical relationship: its lookup to the parent record (e.g. account's
 * parentaccountid), named in the hierarchy value picker's hint (#98). Undefined when there is none.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/query-hierarchical-data
 */
export async function getHierarchicalRelationship(
	logicalName: string,
): Promise<{ schemaName: string; referencingAttribute: string } | undefined> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	const result = await window.dataverseAPI!.queryData(
		`EntityDefinitions(LogicalName='${odataString(logicalName)}')/OneToManyRelationships?$select=SchemaName,ReferencingAttribute&$filter=IsHierarchical eq true`,
	);
	const relationship = result.value[0] as { SchemaName?: string; ReferencingAttribute?: string } | undefined;
	return relationship?.SchemaName && relationship.ReferencingAttribute
		? { schemaName: relationship.SchemaName, referencingAttribute: relationship.ReferencingAttribute }
		: undefined;
}

/** A record found by the hierarchy value picker (#98) */
export interface RecordReference {
	id: string;
	name: string;
}

/** Records whose primary name starts with the text (recordSearch.ts, #98). */
export async function searchRecordsByName(target: RecordSearchTarget, text: string): Promise<RecordReference[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	const result = await window.dataverseAPI!.queryData(buildRecordSearchQuery(target, text));
	return result.value.map((record) => ({
		id: String(record[target.idAttribute] ?? ""),
		name: record[target.nameAttribute] == null ? "" : String(record[target.nameAttribute]),
	}));
}

/** One record's primary name, or undefined when the record can't be read (#98). */
export async function getRecordName(target: RecordSearchTarget, id: string): Promise<string | undefined> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	try {
		const record = (await window.dataverseAPI!.queryData(buildRecordNameQuery(target, id))) as unknown as Record<string, unknown>;
		const name = record[target.nameAttribute];
		return name == null ? "" : String(name);
	} catch (error) {
		debugLog("fetchXmlAPI", `Record ${id} not found for its name:`, error);
		return undefined;
	}
}

/**
 * The total number of records of an entity, unfiltered, from a snapshot within the last 24 hours
 * (RetrieveTotalRecordCount function, #98). Offered when a query's total record count reaches its
 * 5,000 limit. Verified live (2026-10-17): RetrieveTotalRecordCount(EntityNames=@p)?@p=["account"]
 * returns EntityRecordCountCollection { Keys: ["account"], Values: [158965] }.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/retrievetotalrecordcount
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/count-rows?tabs=webapi
 */
export async function retrieveTotalRecordCount(logicalName: string): Promise<number | undefined> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	const result = (await window.dataverseAPI!.queryData(
		`RetrieveTotalRecordCount(EntityNames=@p)?@p=${encodeURIComponent(JSON.stringify([logicalName]))}`,
	)) as unknown as { EntityRecordCountCollection?: { Keys?: string[]; Values?: number[] } };
	return totalFromRecordCountCollection(result.EntityRecordCountCollection, logicalName);
}

/** The count for one entity from RetrieveTotalRecordCountResponse's EntityRecordCountCollection. */
export function totalFromRecordCountCollection(
	collection: { Keys?: string[]; Values?: number[] } | undefined,
	logicalName: string,
): number | undefined {
	const index = collection?.Keys?.findIndex((key) => key.toLowerCase() === logicalName.toLowerCase()) ?? -1;
	const value = index >= 0 ? collection?.Values?.[index] : undefined;
	return typeof value === "number" ? value : undefined;
}

/**
 * Execute FetchXML query with formatted values
 * Formatted values are returned with @OData.Community.Display.V1.FormattedValue annotations
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/fetchxml/retrieve-data?tabs=webapi
 */
export async function executeFetchXml(fetchXml: string): Promise<FetchXmlResult> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("fetchXmlAPI", `📡 Executing FetchXML query...`);
	debugLog("fetchXmlAPI", `FetchXML:\n${fetchXml}`);

	// queryData sends the same request as fetchXmlQuery (GET with Prefer: odata.include-annotations="*"),
	// but with the metadata EntitySetName. fetchXmlQuery is only the fallback when metadata is unavailable.
	let entitySet: Awaited<ReturnType<typeof resolveFetchXmlEntitySet>>;
	try {
		entitySet = await resolveFetchXmlEntitySet(fetchXml);
	} catch (error) {
		console.warn("executeFetchXml: entity set name lookup failed, using fetchXmlQuery:", error);
	}

	const query = (set: string) =>
		window.dataverseAPI!.queryData(`${set}?fetchXml=${encodeURIComponent(fetchXml)}`);
	let response: unknown;
	if (!entitySet) {
		response = await window.dataverseAPI!.fetchXmlQuery(fetchXml);
	} else {
		try {
			response = await query(entitySet.entitySetName);
		} catch (error) {
			// A few system entities (e.g. archivecleanupinfo) report an EntitySetName the service doesn't
			// route; their LogicalCollectionName does (verified live). Retry once with it.
			const fallback = entitySet.logicalCollectionName;
			if (!fallback || fallback === entitySet.entitySetName || !isUnknownSegmentError(error)) throw error;
			response = await query(fallback);
		}
	}
	const result = response as DataverseFetchXmlResponse;

	// Raw response for inspection (debug only: logging every page of a large Retrieve All keeps all
	// rows alive in the console)
	debugLog("fetchXmlAPI", "Raw FetchXML response:", result);

	// The host sends Prefer: odata.include-annotations="*", so formatted values are expected (#71:
	// the old warning about the host's Prefer header is obsolete)
	debugLog("fetchXmlAPI", `Response includes formatted values: ${responseHasFormattedValues(result.value)}`);

	debugLog("fetchXmlAPI", `✅ FetchXML query executed: ${result.value.length} records retrieved`);

	return {
		records: result.value,
		totalRecordCount: result["@Microsoft.Dynamics.CRM.totalrecordcount"],
		totalRecordCountLimitExceeded: result["@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"],
		...pagingFromResponse(result),
	};
}

/**
 * Fetch the next page of a saved view from its @odata.nextLink (#60). Verified live: following the
 * link through queryData returns the next page without overlap and the following nextLink.
 */
export async function executeNextLink(nextLink: string): Promise<FetchXmlResult> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	const path = nextLinkToQueryPath(nextLink);
	if (!path) throw new Error(`Unexpected next page link: ${nextLink}`);
	const result = (await window.dataverseAPI!.queryData(path)) as DataverseFetchXmlResponse;
	return {
		records: result.value,
		totalRecordCount: result["@Microsoft.Dynamics.CRM.totalrecordcount"],
		...pagingFromResponse(result),
	};
}

/**
 * Execute a System View (savedquery) by its ID
 * Uses the optimized ?savedQuery={id} parameter instead of fetching and executing fetchXml
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/retrieve-and-execute-predefined-queries#predefined-queries
 */
export async function executeSystemView(
	entitySetName: string,
	savedQueryId: string,
): Promise<FetchXmlResult> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("fetchXmlAPI", `📡 Executing System View: ${savedQueryId} on ${entitySetName}`);

	// Use queryData with savedQuery parameter (OData query, not FetchXML)
	// GET [Organization URI]/api/data/v9.2/{entitySetName}?savedQuery={savedQueryId}
	const query = `${entitySetName}?savedQuery=${savedQueryId}`;
	const result = (await window.dataverseAPI!.queryData(query)) as DataverseFetchXmlResponse;

	debugLog("fetchXmlAPI", `✅ System View executed: ${result.value.length} records retrieved`);

	// Views page with @odata.nextLink (verified), not the FetchXML paging cookie
	return {
		records: result.value,
		totalRecordCount: undefined,
		...pagingFromResponse(result),
	};
}

/**
 * Execute a Personal View (userquery) by its ID
 * Uses the optimized ?userQuery={id} parameter instead of fetching and executing fetchXml
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/retrieve-and-execute-predefined-queries#predefined-queries
 */
export async function executePersonalView(
	entitySetName: string,
	userQueryId: string,
): Promise<FetchXmlResult> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("fetchXmlAPI", `📡 Executing Personal View: ${userQueryId} on ${entitySetName}`);

	// Use queryData with userQuery parameter (OData query, not FetchXML)
	// GET [Organization URI]/api/data/v9.2/{entitySetName}?userQuery={userQueryId}
	const query = `${entitySetName}?userQuery=${userQueryId}`;
	const result = (await window.dataverseAPI!.queryData(query)) as DataverseFetchXmlResponse;

	debugLog("fetchXmlAPI", `✅ Personal View executed: ${result.value.length} records retrieved`);

	// Views page with @odata.nextLink (verified), not the FetchXML paging cookie
	return {
		records: result.value,
		totalRecordCount: undefined,
		...pagingFromResponse(result),
	};
}

/**
 * Call WhoAmI to get current user context. Memoized per connection (#71): startup called it from
 * four places; resetDataverseSessionCache() clears it when the connection changes. A failed call
 * (null) isn't kept.
 */
export function whoAmI(): Promise<WhoAmIResponse | null> {
	return memoizeForSession("WhoAmI", retrieveWhoAmI, (user) => user !== null);
}

async function retrieveWhoAmI(): Promise<WhoAmIResponse | null> {
	if (!isDataverseAvailable()) {
		console.warn("PPTB Dataverse API not available");
		return null;
	}

	try {
		// Call WhoAmI using execute with proper signature from documentation
		const result = (await window.dataverseAPI!.execute({
			operationName: "WhoAmI",
			operationType: "function",
		})) as unknown as WhoAmIResponse;
		return result;
	} catch (error) {
		console.error("WhoAmI failed:", error);
		return null;
	}
}

/**
 * The user's Dataverse time zone (Personal Options), for date-time condition values (#207) and
 * the Excel export (#184). Memoized per connection; undefined when it can't be read, so callers
 * fall back to values Dataverse reads in the user's time zone anyway.
 */
export function loadUserTimeZone(): Promise<DataverseTimeZone | undefined> {
	return memoizeForSession("userTimeZone", retrieveUserTimeZone, (tz) => tz !== undefined);
}

async function retrieveUserTimeZone(): Promise<DataverseTimeZone | undefined> {
	if (!isDataverseAvailable()) return undefined;
	try {
		const user = await whoAmI();
		if (!user) return undefined;
		const settings = (await window.dataverseAPI!.queryData(
			`usersettingscollection(${user.UserId})?$select=${USER_SETTINGS_TIME_ZONE_COLUMNS.join(",")}`,
		)) as Record<string, unknown>;
		// The name is only for the hint: without it, the rules still apply
		let name = "";
		if (typeof settings.timezonecode === "number") {
			try {
				const definitions = (await window.dataverseAPI!.queryData(
					`timezonedefinitions?$select=standardname&$filter=timezonecode eq ${settings.timezonecode}`,
				)) as { value?: { standardname?: string }[] };
				name = definitions.value?.[0]?.standardname ?? "";
			} catch (error) {
				console.error("Failed to read the time zone's name:", error);
			}
		}
		return timeZoneFromUserSettings(settings, name);
	} catch (error) {
		console.error("Failed to read the user's time zone:", error);
		return undefined;
	}
}

/** Date and money formatting metadata of one attribute, for the Excel export (#184) */
export interface FormattingMetadata {
	Format?: "DateOnly" | "DateAndTime";
	DateTimeBehavior?: { Value?: string };
	Precision?: number;
	/** Money: 0 = Precision, 1 = the organization's pricing precision, 2 = the currency's precision */
	PrecisionSource?: number;
}

/**
 * The formats and behaviours of an entity's date attributes, and the precision of its money,
 * decimal and floating point ones (#184), which the attribute list doesn't include: one request
 * per type. Memoized per connection.
 */
export function loadFormattingMetadata(entityLogicalName: string): Promise<Map<string, FormattingMetadata>> {
	return memoizeForSession(`formatting:${entityLogicalName}`, async () => {
		const result = new Map<string, FormattingMetadata>();
		if (!isDataverseAvailable()) return result;
		const base = `EntityDefinitions(LogicalName='${odataString(entityLogicalName)}')/Attributes`;
		const responses = await Promise.all([
			window.dataverseAPI!.queryData(`${base}/Microsoft.Dynamics.CRM.DateTimeAttributeMetadata?$select=LogicalName,Format,DateTimeBehavior`),
			window.dataverseAPI!.queryData(`${base}/Microsoft.Dynamics.CRM.MoneyAttributeMetadata?$select=LogicalName,Precision,PrecisionSource`),
			window.dataverseAPI!.queryData(`${base}/Microsoft.Dynamics.CRM.DecimalAttributeMetadata?$select=LogicalName,Precision`),
			window.dataverseAPI!.queryData(`${base}/Microsoft.Dynamics.CRM.DoubleAttributeMetadata?$select=LogicalName,Precision`),
		]);
		for (const row of responses.flatMap((response) => (response as { value?: unknown[] }).value ?? [])) {
			const { LogicalName, ...meta } = row as FormattingMetadata & { LogicalName: string };
			result.set(LogicalName, meta);
		}
		return result;
	});
}

/**
 * Convert a FetchXML query to Dataverse SQL using the undocumented FetchXMLToSQL
 * Web API function.
 *
 * ⚠ Note: FetchXMLToSQL is an undocumented/internal API. It may change without notice.
 *
 * @param fetchXml  The FetchXML string to convert.
 * @returns         The SQL string.
 * @throws          If the API call fails or Dataverse is not available.
 */
export async function fetchXmlToSQL(fetchXml: string): Promise<string> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}
	// Escaped and URL-encoded here, not by the host (#176)
	const result = (await window.dataverseAPI!.queryData(
		functionWithString("FetchXMLToSQL", "FetchXml", fetchXml),
	)) as unknown as Record<string, unknown>;

	// The raw result, to debug the undocumented response shape (debug only, #71)
	debugLog("fetchXmlAPI", "FetchXMLToSQL raw result:", result);

	// Try known property names first, then fall back to any non-metadata string value
	const sql =
		(typeof result["value"] === "string" ? result["value"] : undefined) ??
		(typeof result["FetchXMLToSQL"] === "string" ? result["FetchXMLToSQL"] : undefined) ??
		(typeof result["SQL"] === "string" ? result["SQL"] : undefined) ??
		Object.entries(result)
			.filter(([k]) => !k.startsWith("@"))
			.map(([, v]) => v)
			.find((v): v is string => typeof v === "string");

	if (sql == null) {
		const keys = Object.keys(result)
			.filter((k) => !k.startsWith("@"))
			.join(", ");
		console.error("FetchXMLToSQL unexpected response:", result);
		throw new Error(
			`FetchXMLToSQL returned an unexpected response format. Non-metadata keys: [${keys || "none"}]`,
		);
	}
	return sql;
}

/** RetrieveUserPrivilegeByPrivilegeName: whether the user has the privilege (throws on failure). */
async function retrieveUserPrivilege(userId: string, privilegeName: string): Promise<boolean> {
	const query = `systemusers(${userId})/Microsoft.Dynamics.CRM.RetrieveUserPrivilegeByPrivilegeName(PrivilegeName='${odataString(privilegeName)}')`;
	const result = await window.dataverseAPI!.queryData(query);
	// RetrieveUserPrivilegeByPrivilegeName returns the response directly, not wrapped in .value
	const response = result as unknown as PrivilegeCheckResponse;
	const hasPrivilege = !!response?.RolePrivileges?.length;
	debugLog("accessAPI", `checkPrivilegeByName(${privilegeName}): ${hasPrivilege ? "granted" : "denied"}`);
	return hasPrivilege;
}

/**
 * Check if user has a specific privilege using RetrieveUserPrivilegeByPrivilegeName. Memoized per
 * user and privilege for the connection (#71): AppShell, EntitySelector, the record actions and
 * Save view asked for the same privileges; resetDataverseSessionCache() clears it.
 */
export async function checkPrivilegeByName(
	userId: string,
	privilegeName: string,
): Promise<boolean> {
	if (!isDataverseAvailable()) {
		return false;
	}

	try {
		return await memoizeForSession(`privilege:${userId}:${privilegeName}`, () =>
			retrieveUserPrivilege(userId, privilegeName),
		);
	} catch (error) {
		// 0x80040203 = privilege name does not exist in this environment — expected for
		// some entity-specific privileges (e.g. prvDeleteGitorganization). Warn, not error.
		const message = error instanceof Error ? error.message : String(error);
		const isMissingPrivilege = message.includes("0x80040203") || message.includes("does not exist");
		if (isMissingPrivilege) {
			console.warn(
				`checkPrivilegeByName(${privilegeName}): privilege not found in environment, treating as denied.`,
			);
		} else {
			console.error(`checkPrivilegeByName(${privilegeName}) failed:`, error);
		}
		return false;
	}
}

/**
 * Get access summary with privilege checks and mode determination
 */
export async function getAccessSummary(): Promise<AccessSummary | null> {
	const user = await whoAmI();
	if (!user) {
		return null;
	}

	const [canReadPublisher, canReadSolution, canReadCustomization] = await Promise.all([
		checkPrivilegeByName(user.UserId, "prvReadPublisher"),
		checkPrivilegeByName(user.UserId, "prvReadSolution"),
		checkPrivilegeByName(user.UserId, "prvReadCustomization"),
	]);

	return {
		userId: user.UserId,
		canReadPublisher,
		canReadSolution,
		canReadCustomization,
		fullFilterMode: canReadCustomization && canReadSolution && canReadPublisher,
		solutionsOnlyMode: canReadCustomization && canReadSolution && !canReadPublisher,
		publishersOnlyMode: canReadCustomization && canReadPublisher && !canReadSolution,
		metadataOnlyMode: canReadCustomization && !canReadSolution && !canReadPublisher,
		noAccessMode: !canReadCustomization,
	};
}

/**
 * Get publishers with their solutions in one call (Full Filter mode optimization)
 * Uses $expand to retrieve publishers and qualifying solutions together
 * Filters: isreadonly eq false (custom publishers only), isvisible eq true (visible solutions only)
 */
export async function getPublishersWithSolutions(): Promise<PublisherWithSolutions[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	try {
		const query =
			"publishers" +
			"?$select=publisherid,friendlyname,uniquename,customizationprefix" +
			"&$filter=isreadonly eq false and publisher_solution/any(s: s/isvisible eq true and s/solution_solutioncomponent/any(c: c/componenttype eq 1))" +
			"&$expand=publisher_solution($select=solutionid,friendlyname,isvisible,ismanaged,uniquename,version;$filter=isvisible eq true and solution_solutioncomponent/any(c: c/componenttype eq 1))" +
			"&$orderby=friendlyname asc";

		debugLog("publisherAPI", `📡 GET Publishers with Solutions (expanded): ${query}`);

		const result = await window.dataverseAPI!.queryData(query);
		const publisherData = result.value as unknown as Array<
			Publisher & { publisher_solution: Solution[] }
		>;

		const publishersWithSolutions: PublisherWithSolutions[] = publisherData.map((p) => ({
			publisher: {
				publisherid: p.publisherid,
				friendlyname: p.friendlyname,
				uniquename: p.uniquename,
				customizationprefix: p.customizationprefix,
			},
			solutions: p.publisher_solution || [],
		}));

		debugLog(
			"publisherAPI",
			`✅ Publishers with Solutions retrieved: ${
				publishersWithSolutions.length
			} publishers, ${publishersWithSolutions.reduce(
				(sum, p) => sum + p.solutions.length,
				0,
			)} total solutions`,
		);

		return publishersWithSolutions;
	} catch (error) {
		console.error("getPublishersWithSolutions: API call failed:", error);
		throw error;
	}
}

/**
 * Get all publishers
 */
export async function getPublishers(): Promise<Publisher[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	try {
		const query =
			"publishers?$select=publisherid,friendlyname,uniquename,customizationprefix&$filter=isreadonly eq false&$orderby=friendlyname asc";
		debugLog("publisherAPI", `📡 GET Publishers: ${query}`);

		const result = await window.dataverseAPI!.queryData(query);
		const publishers = result.value as unknown as Publisher[];

		debugLog("publisherAPI", `✅ Publishers retrieved: ${publishers.length}`);
		return publishers;
	} catch (error) {
		console.error("getPublishers: API call failed:", error);
		throw error;
	}
}

/**
 * Get solutions filtered by publisher IDs
 * Only returns solutions that contain entities (componenttype=1)
 */
export async function getSolutionsByPublishers(publisherIds: string[]): Promise<Solution[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	if (!publisherIds.length) {
		return [];
	}

	try {
		const SELECT =
			"$select=solutionid,friendlyname,uniquename,solutionpackageversion,_publisherid_value,isvisible,ismanaged";
		const HAS_ENTITIES = "solution_solutioncomponent/any(c: c/componenttype eq 1)";
		const IS_VISIBLE = "isvisible eq true";

		// Chunk publisher IDs to avoid URL length limits
		const chunkSize = 10;
		const chunks: string[][] = [];
		for (let i = 0; i < publisherIds.length; i += chunkSize) {
			chunks.push(publisherIds.slice(i, i + chunkSize));
		}

		const allSolutions = await Promise.all(
			chunks.map(async (chunk) => {
				const publisherFilter = chunk.map((id) => `_publisherid_value eq ${id}`).join(" or ");
				const query = `solutions?${SELECT}&$filter=(${publisherFilter}) and ${IS_VISIBLE} and ${HAS_ENTITIES}&$orderby=friendlyname asc`;

				debugLog("solutionAPI", `📡 GET Solutions for publishers: ${chunk.length} IDs`);
				const result = await window.dataverseAPI!.queryData(query);
				return result.value as unknown as Solution[];
			}),
		);

		// Flatten and dedupe
		const solutionMap = new Map<string, Solution>();
		allSolutions.flat().forEach((s) => solutionMap.set(s.solutionid, s));
		const solutions = Array.from(solutionMap.values());

		debugLog("solutionAPI", `✅ Solutions retrieved: ${solutions.length}`);
		return solutions;
	} catch (error) {
		console.error("getSolutionsByPublishers: API call failed:", error);
		throw error;
	}
}

/**
 * Get all solutions that contain entities (no publisher filter)
 * Filters: isvisible eq true (visible solutions only)
 */
export async function getAllSolutionsWithEntities(): Promise<Solution[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	try {
		const SELECT =
			"$select=solutionid,friendlyname,uniquename,version,_publisherid_value,isvisible,ismanaged";
		const HAS_ENTITIES = "solution_solutioncomponent/any(c: c/componenttype eq 1)";
		const IS_VISIBLE = "isvisible eq true";
		// Exclude solutions from readonly publishers (same logic as getPublishersWithSolutions)
		// to avoid showing Microsoft's internal/system solutions in the list.
		const NON_INTERNAL_PUBLISHER = "publisherid/isreadonly eq false";
		const query = `solutions?${SELECT}&$filter=${IS_VISIBLE} and ${HAS_ENTITIES} and ${NON_INTERNAL_PUBLISHER}&$orderby=friendlyname asc`;

		debugLog("solutionAPI", `📡 GET All Solutions with entities (visible only)`);
		const result = await window.dataverseAPI!.queryData(query);
		const solutions = result.value as unknown as Solution[];

		debugLog("solutionAPI", `✅ Solutions retrieved: ${solutions.length}`);
		return solutions;
	} catch (error) {
		console.error("getAllSolutionsWithEntities: API call failed:", error);
		throw error;
	}
}

/**
 * Get solution components (entities only) for given solution IDs
 * NOTE: Virtual entity msdyn_solutioncomponentsummaries doesn't support OR filters properly,
 * so we query each solution individually and combine results
 */
export async function getSolutionComponents(solutionIds: string[]): Promise<SolutionComponent[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	if (!solutionIds.length) {
		return [];
	}

	try {
		const SELECT =
			"$select=msdyn_name,msdyn_displayname,msdyn_logicalcollectionname,msdyn_solutionid,msdyn_componenttype";
		const COMPONENT_TYPE_ENTITY = 1;

		debugLog("solutionComponentAPI", "Querying solutions individually:", solutionIds);

		// Query each solution individually (virtual entity doesn't support OR filters properly)
		const allComponents = await Promise.all(
			solutionIds.map(async (solutionId) => {
				const query = `msdyn_solutioncomponentsummaries?${SELECT}&$filter=msdyn_componenttype eq ${COMPONENT_TYPE_ENTITY} and msdyn_solutionid eq '${odataString(solutionId)}'`;

				debugLog("solutionComponentAPI", `📡 GET Solution components for solution: ${solutionId}`, query);
				const result = await window.dataverseAPI!.queryData(query);
				debugLog("solutionComponentAPI", `Solution ${solutionId}: ${result.value.length} components`);
				return result.value as unknown as SolutionComponent[];
			}),
		);

		// Flatten and deduplicate by msdyn_name
		const componentMap = new Map<string, SolutionComponent>();
		allComponents.flat().forEach((component) => {
			if (component.msdyn_name && !componentMap.has(component.msdyn_name)) {
				componentMap.set(component.msdyn_name, component);
			}
		});

		const components = Array.from(componentMap.values());

		// Sort by display name
		components.sort((a, b) => {
			const nameA = a.msdyn_displayname || a.msdyn_name || "";
			const nameB = b.msdyn_displayname || b.msdyn_name || "";
			return nameA.localeCompare(nameB);
		});

		debugLog(
			"solutionComponentAPI",
			`✅ Components retrieved: ${components.length} unique entities`,
		);
		return components;
	} catch (error) {
		console.error("getSolutionComponents: API call failed:", error);
		throw error;
	}
}

/**
 * Get every entity's metadata for the Publisher/Solution and Solution scopes (cached, de-duplicated).
 * Not filtered for Advanced Find: the entity picker applies the "Advanced Find only" setting
 * locally in every scope (#57), so toggling it doesn't refetch.
 */
export async function getAllEntitiesForScopes(): Promise<EntityMetadata[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	const cached = metadataCache.getAllEntityMetadata();
	if (cached) return cached;

	const inFlight = metadataCache.getAllEntityMetadataPromise();
	if (inFlight) return inFlight;

	const isCurrent = metadataCache.generationGuard();
	try {
		const promise = getAllEntities(false);
		metadataCache.setAllEntityMetadataPromise(promise);
		const entities = await promise;
		debugLog("metadataAPI", `getAllEntitiesForScopes - Fetched: ${entities.length} entities`);
		// Not after a connection change (#66): these are the previous environment's entities
		if (isCurrent()) {
			metadataCache.setAllEntityMetadata(entities);
			metadataCache.clearAllEntityMetadataPromise();
		}
		return entities;
	} catch (error) {
		if (isCurrent()) metadataCache.clearAllEntityMetadataPromise();
		console.error("[API] getAllEntitiesForScopes - Failed:", error);
		throw error;
	}
}

/**
 * Filter global entity metadata cache by entity logical names
 * Returns entities from cache (instant, no API call)
 */
export function filterCachedEntitiesByNames(logicalNames: string[]): EntityMetadata[] {
	const allEntities = metadataCache.getAllEntityMetadata();
	if (!allEntities) {
		console.warn("[API] filterCachedEntitiesByNames - Global cache not available, returning empty");
		return [];
	}

	const nameSet = new Set(logicalNames);
	const filtered = allEntities.filter((entity: EntityMetadata) => nameSet.has(entity.LogicalName));

	debugLog("metadataAPI", `filterCachedEntitiesByNames: ${filtered.length} of ${logicalNames.length} entities`);

	return filtered;
}

// =============================================================================
// Saved View / User Query APIs
// =============================================================================

/**
 * Represents a saved view (System View or Personal View)
 */
export interface SavedView {
	id: string;
	name: string;
	fetchxml: string;
	layoutxml?: string;
	type: "system" | "personal";
	description?: string;
	isDefault?: boolean;
	/** The querytype column: 0 for public views (savedViews.ts getQueryTypeLabel, #93) */
	queryType?: number;
	/** System views: the iscustomizable managed property; false means it can't be overwritten (#93) */
	isCustomizable?: boolean;
	/** System views: part of a managed solution (ismanaged), so saving adds an unmanaged layer (#142) */
	isManaged?: boolean;
	/** System views: the solution the view belongs to (solutionid), named in the managed warning (#142) */
	solutionId?: string;
	/** Personal views: the owner (ownerid) and its formatted value, for views shared with the user (#142) */
	ownerId?: string;
	ownerName?: string;
}

/**
 * Information about a loaded view for execution optimization
 * Used to determine if the view can be executed via savedQuery/userQuery
 * or if it needs to fall back to fetchXmlQuery execution
 */
export interface LoadedViewInfo {
	/** View ID (savedqueryid or userqueryid) */
	id: string;
	/** View type for determining execution method */
	type: "system" | "personal";
	/** The view's FetchXML as stored in Dataverse; it is compared in generated form once parsed (#93) */
	originalFetchXml: string;
	/** Entity set name for the execution URL */
	entitySetName: string;
	/** View name for display purposes */
	name: string;
	/** Optional LayoutXML for column configuration */
	layoutxml?: string;
	/** The view's querytype, kept when it is overwritten (#93) */
	queryType?: number;
	/** System views: false when the view isn't customizable, so it can't be overwritten (#93) */
	isCustomizable?: boolean;
	/** The view's description, pre-filled when it's overwritten (#142) */
	description?: string;
	/** System views: the default view of its query type (isdefault, #142) */
	isDefault?: boolean;
	/** System views: managed (ismanaged) and its solution, for the overwrite warning (#142) */
	isManaged?: boolean;
	solutionId?: string;
	/** Personal views: the owner, for views shared with the user (#142) */
	ownerId?: string;
	ownerName?: string;
}

// LayoutXML parsing and generation live in model/layoutxml.ts (#81: the copies here were unused)

/** Views with FetchXML to load; a record without any can't be loaded into the tree. */
function hasFetchXml(record: Record<string, unknown>): boolean {
	return typeof record.fetchxml === "string" && record.fetchxml.trim() !== "";
}

/**
 * Get System Views (savedquery) for an entity: every active, published view of every query type.
 * Public views (querytype 0) come first; the others (Advanced Find, Associated, Quick Find,
 * Lookup, Outlook / offline filters and templates...) are listed with a type label rather than
 * hidden (#93, savedViews.ts). The querytype column stores all of them:
 * https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#types-of-views
 * @param entityLogicalName The logical name of the entity (e.g., "account")
 */
export async function getSystemViews(entityLogicalName: string): Promise<SavedView[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 GET System Views for '${entityLogicalName}'`);

	try {
		// Query savedquery entity using queryData for full OData support
		// returnedtypecode: Entity logical name as string
		// componentstate = 0: Published (not unpublished/deleted)
		// statecode = 0: Active views only
		const query =
			`savedqueries?$select=savedqueryid,name,description,fetchxml,layoutxml,isdefault,querytype,iscustomizable,ismanaged,solutionid` +
			`&$filter=(returnedtypecode eq '${odataString(entityLogicalName)}' and componentstate eq 0 and statecode eq 0)` +
			`&$orderby=name asc`;

		debugLog("viewAPI", `Query: ${query}`);

		const result = await window.dataverseAPI!.queryData(query);

		const views: SavedView[] = sortViewsByQueryType(
			(result.value || []).filter(hasFetchXml).map((record) => ({
				id: record.savedqueryid as string,
				name: record.name as string,
				fetchxml: record.fetchxml as string,
				layoutxml: record.layoutxml as string | undefined,
				type: "system" as const,
				description: (record.description as string | null) ?? undefined,
				isDefault: record.isdefault as boolean | undefined,
				queryType: record.querytype as number | undefined,
				isCustomizable: readManagedPropertyValue(record.iscustomizable),
				isManaged: record.ismanaged as boolean | undefined,
				solutionId: (record.solutionid as string | null) ?? undefined,
			})),
		);

		debugLog(
			"viewAPI",
			`✅ System Views retrieved for '${entityLogicalName}': ${views.length} views`,
		);

		return views;
	} catch (error) {
		console.error(`getSystemViews: Failed to get system views for '${entityLogicalName}':`, error);
		throw error;
	}
}

/**
 * Get Personal Views (userquery) for an entity: every active one the user can read. That includes
 * views shared with the user or their teams, not only the ones they own: a user query "can be
 * assigned and shared with other users, and can be viewed by other users depending on the query's
 * access privileges" (#93). Retrieving records already applies those access rights.
 * https://learn.microsoft.com/power-apps/developer/data-platform/saved-queries
 * Like system views, every query type is listed, public views first (savedViews.ts).
 * @param entityLogicalName The logical name of the entity (e.g., "account")
 */
export async function getPersonalViews(entityLogicalName: string): Promise<SavedView[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 GET Personal Views for '${entityLogicalName}'`);

	try {
		// Query userquery entity using queryData for full OData support
		// returnedtypecode: Entity logical name as string
		// statecode = 0: Active views only
		const query =
			`userqueries?$select=userqueryid,name,description,fetchxml,layoutxml,querytype,_ownerid_value` +
			`&$filter=(statecode eq 0 and returnedtypecode eq '${odataString(entityLogicalName)}')` +
			`&$orderby=name asc`;

		debugLog("viewAPI", `Query: ${query}`);

		const result = await window.dataverseAPI!.queryData(query);

		const views: SavedView[] = sortViewsByQueryType(
			(result.value || []).filter(hasFetchXml).map((record) => ({
				id: record.userqueryid as string,
				name: record.name as string,
				fetchxml: record.fetchxml as string,
				layoutxml: record.layoutxml as string | undefined,
				type: "personal" as const,
				description: (record.description as string | null) ?? undefined,
				queryType: record.querytype as number | undefined,
				// The owner's name is the lookup's formatted value: the host asks for all annotations
				// (Prefer: odata.include-annotations="*"), verified live (#142)
				// https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query/select-columns#formatted-values
				ownerId: (record._ownerid_value as string | null) ?? undefined,
				ownerName:
					(record["_ownerid_value@OData.Community.Display.V1.FormattedValue"] as string | undefined) ?? undefined,
			})),
		);

		debugLog(
			"viewAPI",
			`✅ Personal Views retrieved for '${entityLogicalName}': ${views.length} views`,
		);

		return views;
	} catch (error) {
		console.error(
			`getPersonalViews: Failed to get personal views for '${entityLogicalName}':`,
			error,
		);
		throw error;
	}
}

/**
 * Get all views (System + Personal) for an entity
 * @param entityLogicalName The logical name of the entity
 */
export async function getAllViews(
	entityLogicalName: string,
): Promise<{ systemViews: SavedView[]; personalViews: SavedView[] }> {
	const [systemViews, personalViews] = await Promise.all([
		getSystemViews(entityLogicalName),
		getPersonalViews(entityLogicalName),
	]);

	return { systemViews, personalViews };
}

// =============================================================================
// View Save / Update / Validation APIs
// =============================================================================

/**
 * Validator issue from ValidateFetchXmlExpression
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/validatorissue
 */
export interface ValidatorIssue {
	/** Localized message text describing the issue */
	LocalizedMessageText: string;
	/** Severity level: 0 (Low), 1 (Medium), 2 (High), 3 (Critical) */
	Severity: 0 | 1 | 2 | 3;
	/** Type code categorizing the issue */
	TypeCode: number;
	/** Optional additional properties */
	OptionalPropertyBag?: Record<string, string>;
}

/**
 * Response from ValidateFetchXmlExpression function
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/validatefetchxmlexpressionresponse
 */
export interface ValidateFetchXmlExpressionResponse {
	ValidationResults: {
		Helplink?: string;
		Messages: ValidatorIssue[];
	};
}

/**
 * Validate FetchXML for performance issues before save
 * Returns validation messages with severity levels
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/validatefetchxmlexpression
 */
export async function validateFetchXmlExpression(
	fetchXml: string,
): Promise<ValidateFetchXmlExpressionResponse> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 ValidateFetchXmlExpression - Validating FetchXML...`);

	try {
		// FetchXML as a parameter alias, quotes doubled and URL-encoded: an apostrophe in a value
		// (O'Brien) used to end the literal, failing with 0x80060888 (#176)
		const query = functionWithString("ValidateFetchXmlExpression", "FetchXml", fetchXml);

		const result = (await window.dataverseAPI!.queryData(
			query,
		)) as unknown as ValidateFetchXmlExpressionResponse;

		debugLog(
			"viewAPI",
			`✅ ValidateFetchXmlExpression - Complete: ${
				result.ValidationResults?.Messages?.length || 0
			} messages`,
		);

		return result;
	} catch (error) {
		console.error("validateFetchXmlExpression: Failed:", error);
		throw error;
	}
}

/**
 * Validate a system view's FetchXML before it is created or updated (#93). Throws Dataverse's error
 * when the query isn't valid for the query type. The action is bound to the savedquery collection
 * (POST savedqueries/Microsoft.Dynamics.CRM.ValidateSavedQuery) and has no userquery binding, so
 * it is used for system views only.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/validatesavedquery
 */
export async function validateSavedQuery(fetchXml: string, queryType: number = 0): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 ValidateSavedQuery - Validating saved query...`);

	try {
		await window.dataverseAPI!.execute({
			entityName: "savedquery",
			operationName: "ValidateSavedQuery",
			operationType: "action",
			parameters: {
				FetchXml: fetchXml,
				QueryType: queryType,
			},
		});

		debugLog("viewAPI", `✅ ValidateSavedQuery - Validation passed`);
	} catch (error) {
		console.error("validateSavedQuery: Validation failed:", error);
		throw error;
	}
}

/**
 * Check if user has privileges for saving system views (savedquery)
 * Correct privilege names:
 * - prvWriteQuery: Required to create/update saved queries (system views)
 * - prvWriteCustomization: Required to write customizations
 * - prvPublishCustomization: Required to publish customizations
 * @returns Object with privileges for savedquery operations
 */
export async function checkSavedQueryPrivileges(): Promise<{
	canWrite: boolean;
	canPublish: boolean;
}> {
	const user = await whoAmI();
	if (!user) {
		return { canWrite: false, canPublish: false };
	}

	// Try to check privileges, but handle gracefully if the privilege check fails
	let canWriteQuery = false;
	let canWriteCustomization = false;
	let canPublishCustomization = false;

	try {
		// Check for prvWriteQuery privilege (create/update system views)
		canWriteQuery = await checkPrivilegeByName(user.UserId, "prvWriteQuery");
	} catch (error) {
		console.warn("Could not check prvWriteQuery, assuming false:", error);
	}

	try {
		// Check for prvWriteCustomization privilege
		canWriteCustomization = await checkPrivilegeByName(user.UserId, "prvWriteCustomization");
	} catch (error) {
		console.warn("Could not check prvWriteCustomization, assuming false:", error);
	}

	try {
		// Check for prvPublishCustomization privilege
		canPublishCustomization = await checkPrivilegeByName(user.UserId, "prvPublishCustomization");
	} catch (error) {
		console.warn("Could not check prvPublishCustomization, assuming false:", error);
	}

	// User can write system views if they have both prvWriteQuery and prvWriteCustomization
	const canWrite = canWriteQuery && canWriteCustomization;
	// User can publish if they also have prvPublishCustomization
	const canPublish = canWrite && canPublishCustomization;

	debugLog(
		"viewAPI",
		`✅ SavedQuery privileges: writeQuery=${canWriteQuery}, writeCustomization=${canWriteCustomization}, publishCustomization=${canPublishCustomization} → canWrite=${canWrite}, canPublish=${canPublish}`,
	);

	return { canWrite, canPublish };
}

/**
 * Check if user has privileges for saving personal views (userquery)
 * Correct privilege name:
 * - prvWriteUserQuery: Required to create/update user queries (personal views)
 * @returns Whether user can write personal views
 */
export async function checkUserQueryPrivileges(): Promise<{
	canWrite: boolean;
}> {
	const user = await whoAmI();
	if (!user) {
		return { canWrite: false };
	}

	let canWrite = false;

	try {
		// Check for prvWriteUserQuery privilege
		canWrite = await checkPrivilegeByName(user.UserId, "prvWriteUserQuery");
	} catch (error) {
		console.warn("Could not check prvWriteUserQuery, assuming false:", error);
	}

	debugLog("viewAPI", `✅ UserQuery privileges: canWrite=${canWrite}`);

	return { canWrite };
}

/**
 * Whether the current user can write a personal view (#142): RetrievePrincipalAccess, bound to the
 * user, with the userquery as Target. A view shared with the user without write access returns
 * "ReadAccess" only (verified live). Undefined when it can't be determined: saving is then allowed
 * and Dataverse answers with its own error.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/retrieveprincipalaccess
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/use-web-api-functions#pass-record-reference-to-a-function
 */
export async function checkUserQueryWriteAccess(userQueryId: string): Promise<boolean | undefined> {
	if (!isDataverseAvailable()) return undefined;
	try {
		const user = await whoAmI();
		if (!user) return undefined;
		const response = (await window.dataverseAPI!.queryData(
			`systemusers(${user.UserId})/Microsoft.Dynamics.CRM.RetrievePrincipalAccess(Target=@tid)?@tid={'@odata.id':'userqueries(${userQueryId})'}`,
		)) as unknown as { AccessRights?: unknown };
		debugLog("viewAPI", `RetrievePrincipalAccess(userquery ${userQueryId}): ${String(response?.AccessRights)}`);
		return hasWriteAccess(response?.AccessRights);
	} catch (error) {
		console.warn("checkUserQueryWriteAccess: couldn't determine, allowing:", error);
		return undefined;
	}
}

/**
 * Display and unique name of the solution a managed view belongs to (savedquery.solutionid), for
 * the managed view warning (#142). Null when it can't be read. Memoized per connection.
 */
export function getSolutionLabel(solutionId: string): Promise<{ friendlyName: string; uniqueName: string } | null> {
	return memoizeForSession(
		`solutionLabel:${solutionId}`,
		async () => {
			if (!isDataverseAvailable()) return null;
			try {
				const solution = (await window.dataverseAPI!.queryData(
					`solutions(${solutionId})?$select=friendlyname,uniquename`,
				)) as unknown as { friendlyname?: string; uniquename?: string };
				if (!solution?.uniquename) return null;
				return { friendlyName: solution.friendlyname || solution.uniquename, uniqueName: solution.uniquename };
			} catch (error) {
				console.warn(`getSolutionLabel(${solutionId}) failed:`, error);
				return null;
			}
		},
		(label) => label !== null,
	);
}

/**
 * The entity's default public view (savedquery isdefault, querytype 0), named in the "Set as
 * default view" hint (#142). There's one: "You can set only one active public view as the default
 * view" (verified live: one default per query type). Null when there's none or it can't be read.
 * https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#set-a-public-view-as-the-default-view
 */
export async function getDefaultPublicView(entityLogicalName: string): Promise<{ id: string; name: string } | null> {
	if (!isDataverseAvailable()) return null;
	try {
		const result = await window.dataverseAPI!.queryData(
			`savedqueries?$select=savedqueryid,name&$filter=returnedtypecode eq '${odataString(entityLogicalName)}' and querytype eq 0 and isdefault eq true and statecode eq 0&$top=1`,
		);
		const view = result.value?.[0];
		return view ? { id: view.savedqueryid as string, name: view.name as string } : null;
	} catch (error) {
		console.warn(`getDefaultPublicView(${entityLogicalName}) failed:`, error);
		return null;
	}
}

/**
 * Get unmanaged solutions that can receive components
 * Used for the solution picker when saving system views
 */
export async function getUnmanagedSolutions(): Promise<Solution[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 GET Unmanaged Solutions`);

	try {
		// Query unmanaged solutions (ismanaged eq false)
		// Exclude the Default solution (uniquename = 'Default')
		const query =
			"solutions?" +
			"$select=solutionid,friendlyname,uniquename,version,_publisherid_value,isvisible,ismanaged" +
			"&$filter=ismanaged eq false and isvisible eq true and uniquename ne 'Default'" +
			"&$orderby=friendlyname asc";

		const result = await window.dataverseAPI!.queryData(query);
		const solutions = result.value as unknown as Solution[];

		debugLog("viewAPI", `✅ Unmanaged Solutions retrieved: ${solutions.length}`);
		return solutions;
	} catch (error) {
		console.error("getUnmanagedSolutions: Failed:", error);
		throw error;
	}
}

/**
 * Data for creating a new system view (savedquery)
 */
export interface CreateSavedQueryData {
	name: string;
	fetchxml: string;
	layoutxml: string;
	returnedtypecode: string; // Entity logical name
	/** On update, null clears it (#142) */
	description?: string | null;
	querytype?: number; // 0 = Main/Public view
	/** Public views: make it the entity's default view (#142, savedViews.ts canSetDefaultView) */
	isdefault?: boolean;
}

/**
 * Data for creating a new personal view (userquery)
 */
export interface CreateUserQueryData {
	name: string;
	fetchxml: string;
	layoutxml: string;
	returnedtypecode: string; // Entity logical name
	/** On update, null clears it (#142) */
	description?: string | null;
	querytype?: number; // 0 = Saved view
}

/**
 * Create a new System View (savedquery)
 * @returns The ID of the created view
 */
export async function createSavedQuery(data: CreateSavedQueryData): Promise<string> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 CREATE SavedQuery: "${data.name}"`);

	try {
		const record = {
			name: data.name,
			fetchxml: data.fetchxml,
			layoutxml: data.layoutxml,
			returnedtypecode: data.returnedtypecode,
			description: data.description || null,
			querytype: data.querytype ?? 0,
			...(data.isdefault ? { isdefault: true } : {}),
		};

		const result = await window.dataverseAPI!.create("savedquery", record);

		debugLog("viewAPI", `✅ SavedQuery created: ${result.id}`);
		return result.id;
	} catch (error) {
		console.error("createSavedQuery: Failed:", error);
		throw error;
	}
}

/**
 * Update an existing System View (savedquery)
 */
export async function updateSavedQuery(
	savedQueryId: string,
	data: Partial<CreateSavedQueryData>,
): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 UPDATE SavedQuery: ${savedQueryId}`);

	try {
		const record: Record<string, unknown> = {};
		if (data.name !== undefined) record.name = data.name;
		if (data.fetchxml !== undefined) record.fetchxml = data.fetchxml;
		if (data.layoutxml !== undefined) record.layoutxml = data.layoutxml;
		if (data.description !== undefined) record.description = data.description;
		if (data.isdefault !== undefined) record.isdefault = data.isdefault;

		await window.dataverseAPI!.update("savedquery", savedQueryId, record);

		debugLog("viewAPI", `✅ SavedQuery updated: ${savedQueryId}`);
	} catch (error) {
		console.error("updateSavedQuery: Failed:", error);
		throw error;
	}
}

/**
 * Create a new Personal View (userquery)
 * @returns The ID of the created view
 */
export async function createUserQuery(data: CreateUserQueryData): Promise<string> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 CREATE UserQuery: "${data.name}"`);

	try {
		const record = {
			name: data.name,
			fetchxml: data.fetchxml,
			layoutxml: data.layoutxml,
			returnedtypecode: data.returnedtypecode,
			description: data.description || null,
			querytype: data.querytype ?? 0,
		};

		const result = await window.dataverseAPI!.create("userquery", record);

		debugLog("viewAPI", `✅ UserQuery created: ${result.id}`);
		return result.id;
	} catch (error) {
		console.error("createUserQuery: Failed:", error);
		throw error;
	}
}

/**
 * Update an existing Personal View (userquery)
 */
export async function updateUserQuery(
	userQueryId: string,
	data: Partial<CreateUserQueryData>,
): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 UPDATE UserQuery: ${userQueryId}`);

	try {
		const record: Record<string, unknown> = {};
		if (data.name !== undefined) record.name = data.name;
		if (data.fetchxml !== undefined) record.fetchxml = data.fetchxml;
		if (data.layoutxml !== undefined) record.layoutxml = data.layoutxml;
		if (data.description !== undefined) record.description = data.description;

		await window.dataverseAPI!.update("userquery", userQueryId, record);

		debugLog("viewAPI", `✅ UserQuery updated: ${userQueryId}`);
	} catch (error) {
		console.error("updateUserQuery: Failed:", error);
		throw error;
	}
}

/**
 * Add a component to a solution
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/addsolutioncomponent
 */
export async function addSolutionComponent(
	componentId: string,
	componentType: number,
	solutionUniqueName: string,
	addRequiredComponents: boolean = false,
): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog(
		"viewAPI",
		`📡 AddSolutionComponent: ${componentId} (type=${componentType}) to "${solutionUniqueName}"`,
	);

	try {
		await window.dataverseAPI!.execute({
			operationName: "AddSolutionComponent",
			operationType: "action",
			parameters: {
				ComponentId: componentId,
				ComponentType: componentType,
				SolutionUniqueName: solutionUniqueName,
				AddRequiredComponents: addRequiredComponents,
			},
		});

		debugLog("viewAPI", `✅ Component added to solution: ${componentId} → ${solutionUniqueName}`);
	} catch (error) {
		console.error("addSolutionComponent: Failed:", error);
		throw error;
	}
}

/**
 * Component type for savedquery (system views)
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/solutioncomponent
 */
export const COMPONENT_TYPE_SAVEDQUERY = 26;

/**
 * Publish a system view by publishing its entity (#93 h). PublishXml publishes entities, ribbons,
 * dashboards, option sets, the site map and web resources; a savedquery node isn't documented, and
 * publishing the entity includes its views (and any other pending customizations of that entity).
 * @see https://learn.microsoft.com/dotnet/api/microsoft.crm.sdk.messages.publishxmlrequest.parameterxml
 * @see https://learn.microsoft.com/power-apps/developer/model-driven-apps/publish-request-schema
 */
export async function publishEntityViews(entityLogicalName: string): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 PublishXml for entity: ${entityLogicalName}`);

	try {
		const name = entityLogicalName.replace(/[^a-z0-9_]/gi, "");
		const parameterXml = `<importexportxml><entities><entity>${name}</entity></entities></importexportxml>`;

		await window.dataverseAPI!.execute({
			operationName: "PublishXml",
			operationType: "action",
			parameters: {
				ParameterXml: parameterXml,
			},
		});

		debugLog("viewAPI", `✅ Entity published: ${name}`);
	} catch (error) {
		console.error("publishEntityViews: Failed:", error);
		throw error;
	}
}

/**
 * Check if a solution is managed by unique name
 * @returns true if solution is managed, false if unmanaged, null if not found
 */
export async function isSolutionManaged(solutionUniqueName: string): Promise<boolean | null> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 Check if solution "${solutionUniqueName}" is managed`);

	try {
		const query = `solutions?$select=solutionid,ismanaged&$filter=uniquename eq '${odataString(solutionUniqueName)}'`;
		const result = await window.dataverseAPI!.queryData(query);

		if (!result.value || result.value.length === 0) {
			debugLog("viewAPI", `⚠️ Solution "${solutionUniqueName}" not found`);
			return null;
		}

		const isManaged = result.value[0].ismanaged as boolean;
		debugLog("viewAPI", `✅ Solution "${solutionUniqueName}" ismanaged=${isManaged}`);
		return isManaged;
	} catch (error) {
		console.error("isSolutionManaged: Failed:", error);
		throw error;
	}
}

/**
 * Get solution ID by unique name
 * @returns Solution ID or null if not found
 */
export async function getSolutionIdByUniqueName(
	solutionUniqueName: string,
): Promise<string | null> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 Get solution ID for "${solutionUniqueName}"`);

	try {
		const query = `solutions?$select=solutionid&$filter=uniquename eq '${odataString(solutionUniqueName)}'`;
		const result = await window.dataverseAPI!.queryData(query);

		if (!result.value || result.value.length === 0) {
			debugLog("viewAPI", `⚠️ Solution "${solutionUniqueName}" not found`);
			return null;
		}

		const solutionId = result.value[0].solutionid as string;
		debugLog("viewAPI", `✅ Solution ID: ${solutionId}`);
		return solutionId;
	} catch (error) {
		console.error("getSolutionIdByUniqueName: Failed:", error);
		throw error;
	}
}

/**
 * Export to Excel using Dataverse ExportToExcel action
 * Requires a saved view (system or personal)
 *
 * The ExportToExcel action expects:
 * - View: Object with @odata.type and view ID (savedqueryid or userqueryid)
 * - FetchXml: The FetchXML query string
 * - LayoutXml: The LayoutXML for column configuration
 * - QueryApi: Empty string
 * - QueryParameters: Object with Arguments
 *
 * @param viewId - The savedqueryid (system) or userqueryid (personal)
 * @param viewType - Whether it's a system or personal view
 * @param fetchXml - The FetchXML query
 * @param layoutXml - The LayoutXML for column configuration
 * @param viewName - Name of the view for the filename
 * @returns Base64 encoded Excel file data
 */
export async function exportToExcel(
	viewId: string,
	viewType: "system" | "personal",
	fetchXml: string,
	layoutXml: string,
	viewName: string,
): Promise<{ excelFile: string; filename: string }> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// Build the View object based on view type
	const viewObject =
		viewType === "system"
			? {
					"@odata.type": "Microsoft.Dynamics.CRM.savedquery",
					savedqueryid: viewId,
				}
			: {
					"@odata.type": "Microsoft.Dynamics.CRM.userquery",
					userqueryid: viewId,
				};

	debugLog("exportAPI", `📡 ExportToExcel: ${viewType} view ${viewId}`);
	debugLog("exportAPI", `📡 View object:`, viewObject);
	debugLog(
		"exportAPI",
		`📡 FetchXML length: ${fetchXml.length}, LayoutXML length: ${layoutXml.length}`,
	);

	try {
		// Build the ExportToExcel request matching Dataverse format
		const result = (await window.dataverseAPI!.execute({
			operationName: "ExportToExcel",
			operationType: "action",
			parameters: {
				View: viewObject,
				FetchXml: fetchXml,
				LayoutXml: layoutXml,
				QueryApi: "",
				QueryParameters: {
					Arguments: {
						Count: 0,
						IsReadOnly: true,
						Keys: [],
						Values: [],
					},
				},
			},
		})) as { ExcelFile: string };

		// Generate filename with view name and timestamp
		const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
		const filename = `${viewName} - ${timestamp}.xlsx`;

		debugLog(
			"exportAPI",
			`✅ ExportToExcel successful, file size: ${result.ExcelFile?.length || 0} bytes`,
		);

		return {
			excelFile: result.ExcelFile,
			filename,
		};
	} catch (error) {
		console.error("exportToExcel: Failed:", error);
		throw error;
	}
}

/**
 * Download a Base64 encoded file
 * @param base64Data - The Base64 encoded file data
 * @param filename - The filename for the download
 * @param mimeType - The MIME type of the file
 */
export async function downloadBase64File(
	base64Data: string,
	filename: string,
	mimeType: string = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
): Promise<void> {
	try {
		// Decode Base64 to binary buffer
		const byteCharacters = atob(base64Data);
		const byteArray = new Uint8Array(byteCharacters.length);
		for (let i = 0; i < byteCharacters.length; i++) {
			byteArray[i] = byteCharacters.charCodeAt(i);
		}

		// Use PPTB fileSystem.saveFile when running in the desktop host (opens native save dialog)
		if (window.toolboxAPI?.fileSystem?.saveFile) {
			const saved = await window.toolboxAPI.fileSystem.saveFile(filename, byteArray);
			if (saved) {
				debugLog("exportAPI", `✅ File saved via fileSystem: ${saved}`);
			}
			return;
		}

		// Fallback: browser blob download (standalone / dev context)
		const blob = new Blob([byteArray], { type: mimeType });
		const url = URL.createObjectURL(blob);
		const link = document.createElement("a");
		link.href = url;
		link.download = filename;
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);
		URL.revokeObjectURL(url);

		debugLog("exportAPI", `✅ File downloaded: ${filename}`);
	} catch (error) {
		console.error("downloadBase64File: Failed:", error);
		throw error;
	}
}

// ============================================================================
// RECORD ACTIONS
// ============================================================================

/**
 * Get the Dataverse environment URL from the active connection
 */
export async function getEnvironmentUrl(): Promise<string | null> {
	try {
		// Slot-aware on PPTB ≥ 1.2.7 (#70)
		const conn = await getPrimaryConnection();
		return conn?.url || null;
	} catch (error) {
		console.error("getEnvironmentUrl: Failed:", error);
		return null;
	}
}

/**
 * Build a record URL for opening in browser
 * @param entityName Logical name of the entity
 * @param recordId GUID of the record
 * @param environmentUrl Base URL of the environment
 */
export function buildRecordUrl(
	entityName: string,
	recordId: string,
	environmentUrl: string,
): string {
	// Remove trailing slash from environmentUrl if present to avoid double slashes
	const baseUrl = environmentUrl.endsWith("/") ? environmentUrl.slice(0, -1) : environmentUrl;
	// Standard Dynamics 365 record URL format
	return `${baseUrl}/main.aspx?etn=${entityName}&id=${recordId}&pagetype=entityrecord`;
}

/**
 * The entity set to delete from with executeBatch (#61), or undefined to use the host's
 * dataverseAPI.delete: on hosts without executeBatch (< 1.2.7), or when metadata has no
 * EntitySetName. dataverseAPI.delete takes a logical name and pluralizes it on the host.
 */
async function deleteEntitySet(entityLogicalName: string): Promise<string | undefined> {
	if (!canExecuteBatch()) return undefined;
	try {
		return (await resolveEntitySet(entityLogicalName))?.entitySetName;
	} catch (error) {
		console.warn("Delete: entity set name lookup failed, using the host's delete:", error);
		return undefined;
	}
}

const hostExecuteBatch: ExecuteBatch = (requests) => window.dataverseAPI!.executeBatch(requests);

/**
 * Delete a single record. Uses the metadata EntitySetName through executeBatch (#61).
 * @param entityLogicalName Logical name of the entity (e.g., "account")
 * @param recordId GUID of the record to delete
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/update-delete-entities-using-web-api#basic-delete
 */
export async function deleteRecord(
	entityLogicalName: string,
	recordId: string,
	/** Asked before each attempt and retry (#175): false stops, e.g. after a connection change */
	shouldContinue?: () => boolean | Promise<boolean>,
): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("recordAPI", `📡 Deleting record: ${entityLogicalName}(${recordId})`);

	try {
		const entitySetName = await deleteEntitySet(entityLogicalName);
		if (entitySetName) {
			// The batch path retries service protection errors itself (batchDelete.ts)
			const result = await deleteRecordsInBatches(entitySetName, [recordId], { executeBatch: hostExecuteBatch, shouldContinue });
			if (result.failed > 0) throw new Error(result.errors[0]?.replace(`${recordId}: `, "") || "Delete failed");
		} else {
			// Retried after service protection errors (#96): Dataverse rejected the request unprocessed
			if (shouldContinue && !(await shouldContinue())) throw new Error(CONNECTION_CHANGED_MESSAGE);
			await withThrottleRetry(() => window.dataverseAPI!.delete(entityLogicalName, recordId), { shouldContinue });
		}
		debugLog("recordAPI", `✅ Record deleted: ${entityLogicalName}(${recordId})`);
	} catch (error) {
		console.error("deleteRecord: Failed:", error);
		throw error;
	}
}

/**
 * Check if user has privilege for bulk delete
 */
export async function checkBulkDeletePrivilege(): Promise<boolean> {
	try {
		const user = await whoAmI();
		if (!user) return false;
		return await checkPrivilegeByName(user.UserId, "prvBulkDelete");
	} catch (error) {
		console.error("checkBulkDeletePrivilege: Failed:", error);
		return false;
	}
}

/**
 * Convert FetchXML to QueryExpression using the FetchXmlToQueryExpression function. The FetchXML is
 * escaped and URL-encoded here (functionWithString), not by the host: PPTB Desktop's execute()
 * escapes it, the VS Code extension doesn't (#176, #173).
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/fetchxmltoqueryexpression
 */
export async function fetchXmlToQueryExpression(
	fetchXml: string,
): Promise<Record<string, unknown>> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("recordAPI", `📡 Converting FetchXML to QueryExpression`);

	const result = (await window.dataverseAPI!.queryData(
		functionWithString("FetchXmlToQueryExpression", "FetchXml", fetchXml),
	)) as unknown as { Query: Record<string, unknown> };

	debugLog("recordAPI", `✅ Converted FetchXML to QueryExpression`, result);
	return result.Query ?? (result as unknown as Record<string, unknown>);
}

/**
 * Get the bulkdeleteoperationid from an asyncoperationid
 * The BulkDelete action returns an asyncoperationid, but we need the bulkdeleteoperationid
 * to link to the correct record in the UI
 * @param asyncOperationId The asyncoperationid returned from BulkDelete action
 * @returns The bulkdeleteoperationid or empty string if not found
 */
async function getBulkDeleteOperationId(asyncOperationId: string): Promise<string> {
	if (!asyncOperationId || !isDataverseAvailable()) {
		return "";
	}

	try {
		// Query bulkdeleteoperations to find the one with this asyncoperationid
		const query = `bulkdeleteoperations?$select=bulkdeleteoperationid&$filter=_asyncoperationid_value eq ${asyncOperationId}`;
		const result = await window.dataverseAPI!.queryData(query);

		const operations = result.value as Array<{ bulkdeleteoperationid?: string }>;
		if (operations && operations.length > 0 && operations[0].bulkdeleteoperationid) {
			debugLog(
				"recordAPI",
				`✅ Found bulkdeleteoperationid: ${operations[0].bulkdeleteoperationid}`,
			);
			return operations[0].bulkdeleteoperationid;
		}

		debugLog(
			"recordAPI",
			`⚠️ No bulkdeleteoperationid found for asyncoperationid: ${asyncOperationId}`,
		);
		return "";
	} catch (error) {
		console.error("getBulkDeleteOperationId: Failed:", error);
		return "";
	}
}

/**
 * Submit a bulk delete job for selected records. The query is built here (#177: converting an
 * ID list through FetchXmlToQueryExpression put it all in a GET URL, which failed above ~530 records).
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/delete-data-bulk
 * @param entityLogicalName Logical name of the entity
 * @param primaryIdAttribute Primary ID attribute name (e.g., "accountid")
 * @param recordIds Array of record GUIDs to delete
 * @param jobName Name for the bulk delete job
 * @returns AsyncOperationId and BulkDeleteOperationId for tracking the job
 */
export async function submitBulkDelete(
	entityLogicalName: string,
	primaryIdAttribute: string,
	recordIds: string[],
	jobName: string,
): Promise<{ asyncOperationId: string; bulkDeleteOperationId: string; jobUrl: string }> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog(
		"recordAPI",
		`📡 Submitting bulk delete job: ${jobName} for ${recordIds.length} records`,
	);

	try {
		// primaryIdAttribute in (the selected IDs), with the IDs as strings (#224, #177)
		const queryExpression = selectedRecordsQuery(entityLogicalName, primaryIdAttribute, recordIds);

		// Submit BulkDelete with the QueryExpression
		const result = (await window.dataverseAPI!.execute({
			operationName: "BulkDelete",
			operationType: "action",
			parameters: {
				QuerySet: [queryExpression],
				JobName: jobName,
				SendEmailNotification: false,
				ToRecipients: [],
				CCRecipients: [],
				RecurrencePattern: "",
				StartDateTime: new Date().toISOString(),
			},
		})) as { JobId?: string; AsyncOperationId?: string };

		const asyncOpId = result.AsyncOperationId || result.JobId || "";

		// Query for the bulkdeleteoperationid using the asyncoperationid
		const bulkDeleteOpId = await getBulkDeleteOperationId(asyncOpId);

		// Build job tracking URL using bulkdeleteoperation entity
		const envUrl = await getEnvironmentUrl();
		const baseUrl = envUrl?.endsWith("/") ? envUrl.slice(0, -1) : envUrl;
		const jobUrl =
			baseUrl && bulkDeleteOpId
				? `${baseUrl}/main.aspx?pagetype=entityrecord&etn=bulkdeleteoperation&id=${bulkDeleteOpId}`
				: "";

		debugLog(
			"recordAPI",
			`✅ Bulk delete job submitted: asyncOpId=${asyncOpId}, bulkDeleteOpId=${bulkDeleteOpId}`,
		);

		return {
			asyncOperationId: asyncOpId,
			bulkDeleteOperationId: bulkDeleteOpId,
			jobUrl,
		};
	} catch (error) {
		console.error("submitBulkDelete: Failed:", error);
		throw error;
	}
}

// ============================================================================
// WORKFLOW ACTIONS
// ============================================================================

export interface WorkflowInfo {
	workflowid: string;
	name: string;
	description?: string;
	primaryentity: string;
	category: number; // 0 = Workflow, 2 = Business Rule, etc.
	type: number; // 1 = Definition, 2 = Activation, 3 = Template
	statuscode: number;
}

/**
 * Check if user has privileges to read and execute workflows
 * Returns true if user can both read workflows and execute them
 */
export async function checkWorkflowPrivileges(): Promise<boolean> {
	try {
		const user = await whoAmI();
		if (!user) return false;

		const [canRead, canExecute] = await Promise.all([
			checkPrivilegeByName(user.UserId, "prvReadWorkflow"),
			checkPrivilegeByName(user.UserId, "prvWorkflowExecution"),
		]);

		return canRead && canExecute;
	} catch (error) {
		console.error("checkWorkflowPrivileges: Failed:", error);
		return false;
	}
}

/**
 * Get available on-demand workflows for an entity
 * @param entityLogicalName Logical name of the entity
 * @returns List of workflows that can be run on-demand
 */
export async function getOnDemandWorkflows(entityLogicalName: string): Promise<WorkflowInfo[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("workflowAPI", `📡 Getting on-demand workflows for: ${entityLogicalName}`);

	try {
		// Query for active, on-demand workflows for this entity using OData
		// statecode=1 (Activated), statuscode=2 (Activated), ondemand=true, category=0 (Workflow), type=1 (Definition)
		// primaryentity must be quoted as a string value
		const odataQuery = `workflows?$select=workflowid,name,description,primaryentity,category,type,statuscode&$filter=(statecode eq 1 and statuscode eq 2 and ondemand eq true and category eq 0 and type eq 1 and primaryentity eq '${odataString(entityLogicalName)}')&$orderby=name asc`;

		const result = await window.dataverseAPI!.queryData(odataQuery);

		// Map the records to WorkflowInfo type
		const workflows: WorkflowInfo[] = (result.value || []).map((record) => ({
			workflowid: String(record.workflowid || ""),
			name: String(record.name || ""),
			description: record.description ? String(record.description) : undefined,
			primaryentity: String(record.primaryentity || ""),
			category: Number(record.category || 0),
			type: Number(record.type || 0),
			statuscode: Number(record.statuscode || 0),
		}));

		debugLog("workflowAPI", `✅ Found ${workflows.length} on-demand workflows`);

		return workflows;
	} catch (error) {
		console.error("getOnDemandWorkflows: Failed:", error);
		throw error;
	}
}

/**
 * Execute a workflow on a specific record
 * @param workflowId GUID of the workflow to execute
 * @param recordId GUID of the record to run the workflow on
 * @param entityLogicalName Logical name of the entity (for the target EntityReference)
 */
export async function executeWorkflow(
	workflowId: string,
	recordId: string,
	entityLogicalName: string,
	/** Asked after each throttling wait (#175): false stops instead of retrying */
	shouldContinue?: () => boolean | Promise<boolean>,
): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog(
		"workflowAPI",
		`📡 Executing workflow ${workflowId} on ${entityLogicalName}(${recordId})`,
	);

	try {
		// Retried after service protection errors (#96): Dataverse rejected the request unprocessed
		await withThrottleRetry(() =>
			window.dataverseAPI!.execute({
				operationName: "ExecuteWorkflow",
				operationType: "action",
				entityName: "workflow",
				entityId: workflowId,
				parameters: {
					EntityId: recordId,
				},
			}),
			{ shouldContinue }, // not retried after a connection change (#175)
		);

		debugLog("workflowAPI", `✅ Workflow executed successfully`);
	} catch (error) {
		console.error("executeWorkflow: Failed:", error);
		throw error;
	}
}

/**
 * Progress info for workflow batch execution
 */
export interface WorkflowBatchProgress {
	completed: number;
	total: number;
	succeeded: number;
	failed: number;
	batchesCompleted: number;
	totalBatches: number;
	estimatedSecondsRemaining?: number;
}

/**
 * Execute a workflow on multiple records using $batch requests for parallelization
 * Uses parallel execution via PPTB dataverseAPI for performance
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/execute-batch-operations-using-web-api
 * @param workflowId GUID of the workflow to execute
 * @param recordIds Array of record GUIDs to run the workflow on
 * @param entityLogicalName Logical name of the entity (unused but kept for API compatibility)
 * @param onProgress Callback for progress updates with batch info and ETA
 * @param batchSize Number of concurrent requests per batch (default: 10)
 */
export async function executeWorkflowBatch(
	workflowId: string,
	recordIds: string[],
	entityLogicalName: string,
	onProgress?: (progress: WorkflowBatchProgress) => void,
	batchSize: number = 10,
): Promise<{ succeeded: number; failed: number; errors: string[] }> {
	const result = { succeeded: 0, failed: 0, errors: [] as string[] };
	const total = recordIds.length;
	let completed = 0;

	// Split recordIds into batches of batchSize (default 10) for parallel execution
	const batches: string[][] = [];
	for (let i = 0; i < recordIds.length; i += batchSize) {
		batches.push(recordIds.slice(i, i + batchSize));
	}

	debugLog(
		"workflowAPI",
		`📡 Executing workflow ${workflowId} on ${recordIds.length} records in ${batches.length} batches`,
	);

	// Track timing for ETA calculation
	const batchTimes: number[] = [];
	let batchesCompleted = 0;
	// Stop if the tool's connection changes in place (PPTB ≥ 1.2.7, #66): later requests would run
	// the workflow in the new environment, where the same record IDs can exist
	const stillOnConnection = await connectionGuard();

	// Process batches sequentially (each batch processes in parallel)
	for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
		const batchRecordIds = batches[batchIndex];
		const batchStartTime = Date.now();

		if (!(await stillOnConnection())) {
			const remaining = batches.slice(batchIndex).flat();
			result.failed += remaining.length;
			result.errors.push(...remaining.map((id) => `Record ${id}: Not run: the tool's connection changed`));
			break;
		}

		debugLog(
			"workflowAPI",
			`📡 Executing batch ${batchIndex + 1}/${batches.length} with ${batchRecordIds.length} records`,
		);

		// Execute workflows in parallel within the batch
		const workflowPromises = batchRecordIds.map(async (recordId) => {
			try {
				await executeWorkflow(workflowId, recordId, entityLogicalName, stillOnConnection);
				return { success: true, recordId };
			} catch (error) {
				return {
					success: false,
					recordId,
					error: formatDataverseError(error), // actionable text for documented codes (#96)
				};
			}
		});

		const batchResults = await Promise.all(workflowPromises);
		const batchEndTime = Date.now();
		const batchTime = batchEndTime - batchStartTime;
		batchTimes.push(batchTime);

		// Aggregate results
		for (const execResult of batchResults) {
			if (execResult.success) {
				result.succeeded++;
			} else {
				result.failed++;
				result.errors.push(`Record ${execResult.recordId}: ${execResult.error}`);
			}
			completed++;
		}
		batchesCompleted++;

		// Calculate ETA based on average batch time
		let estimatedSecondsRemaining: number | undefined;
		if (batchTimes.length >= 2) {
			const avgBatchTime = batchTimes.reduce((a, b) => a + b, 0) / batchTimes.length;
			const remainingBatches = batches.length - batchesCompleted;
			estimatedSecondsRemaining = Math.ceil((avgBatchTime * remainingBatches) / 1000);
		}

		onProgress?.({
			completed,
			total,
			succeeded: result.succeeded,
			failed: result.failed,
			batchesCompleted,
			totalBatches: batches.length,
			estimatedSecondsRemaining,
		});
	}

	debugLog(
		"workflowAPI",
		`✅ Workflow execution complete: ${result.succeeded} succeeded, ${result.failed} failed`,
	);
	return result;
}

/**
 * Check if user has delete privilege for a specific entity
 */
export async function checkDeletePrivilege(entityLogicalName: string): Promise<boolean> {
	try {
		const user = await whoAmI();
		if (!user) return false;

		// Entity-specific delete privilege (e.g., prvDeleteAccount)
		const deletePriv = `prvDelete${entityLogicalName
			.charAt(0)
			.toUpperCase()}${entityLogicalName.slice(1)}`;
		return await checkPrivilegeByName(user.UserId, deletePriv);
	} catch (error) {
		console.error("checkDeletePrivilege: Failed:", error);
		return false;
	}
}

// ============================================================================
// BATCH DELETE OPERATIONS
// ============================================================================

export interface BatchDeleteProgress {
	completed: number;
	total: number;
	succeeded: number;
	failed: number;
	batchesCompleted: number;
	totalBatches: number;
	estimatedSecondsRemaining?: number;
}

export interface BatchDeleteResult {
	succeeded: number;
	failed: number;
	errors: string[];
}

/**
 * Delete multiple records. Recommended for 1-100 records; for larger sets, use submitBulkDelete.
 *
 * PPTB ≥ 1.2.7: OData $batch requests through dataverseAPI.executeBatch with the metadata
 * EntitySetName (#61, batchDelete.ts). Older hosts: groups of parallel single DELETE requests
 * through dataverseAPI.delete (host-pluralized entity set; not $batch).
 *
 * @param entityLogicalName Logical name of the entity (e.g., "account")
 * @param recordIds Array of record GUIDs to delete
 * @param onProgress Callback for progress updates with ETA
 * @param batchSize Parallel requests per group on the fallback path (default: 10)
 */
export async function deleteRecordsBatch(
	entityLogicalName: string,
	recordIds: string[],
	onProgress?: (progress: BatchDeleteProgress) => void,
	batchSize: number = 10,
): Promise<BatchDeleteResult> {
	const entitySetName = await deleteEntitySet(entityLogicalName);
	if (entitySetName) {
		debugLog("recordAPI", `📡 Deleting ${recordIds.length} records from ${entitySetName} with $batch`);
		const batchResult = await deleteRecordsInBatches(entitySetName, recordIds, {
			executeBatch: hostExecuteBatch,
			onProgress,
			// Stop if the tool's connection changes in place (#66)
			shouldContinue: await connectionGuard(),
		});
		debugLog("recordAPI", `✅ Delete complete: ${batchResult.succeeded} succeeded, ${batchResult.failed} failed`);
		return batchResult;
	}

	const result: BatchDeleteResult = { succeeded: 0, failed: 0, errors: [] };
	const total = recordIds.length;

	// Split recordIds into batches for parallel processing
	const batches: string[][] = [];
	for (let i = 0; i < recordIds.length; i += batchSize) {
		batches.push(recordIds.slice(i, i + batchSize));
	}

	debugLog("recordAPI", `📡 Deleting ${recordIds.length} records in ${batches.length} batches`);

	// Track timing for ETA calculation
	const batchTimes: number[] = [];
	let completed = 0;
	let batchesCompleted = 0;
	// Older hosts without executeBatch: the same connection guard as the $batch path (#175, #66)
	const stillOnConnection = await connectionGuard();

	// Process batches sequentially (each batch processes in parallel)
	for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
		const batchRecordIds = batches[batchIndex];
		const batchStartTime = Date.now();

		if (!(await stillOnConnection())) {
			const remaining = batches.slice(batchIndex).flat();
			result.failed += remaining.length;
			result.errors.push(...remaining.map((id) => `${id}: ${CONNECTION_CHANGED_MESSAGE}`));
			break;
		}

		// Execute deletes in parallel within the batch
		const deletePromises = batchRecordIds.map(async (recordId) => {
			try {
				// Retried after service protection errors (#96), like deleteRecord's fallback; not after
				// a connection change (#175)
				await withThrottleRetry(() => window.dataverseAPI!.delete(entityLogicalName, recordId), { shouldContinue: stillOnConnection });
				return { success: true, recordId };
			} catch (error) {
				return {
					success: false,
					recordId,
					error: formatDataverseError(error), // actionable text for documented codes (#96)
				};
			}
		});

		const batchResults = await Promise.all(deletePromises);
		const batchEndTime = Date.now();
		const batchTime = batchEndTime - batchStartTime;
		batchTimes.push(batchTime);

		// Aggregate results
		for (const deleteResult of batchResults) {
			if (deleteResult.success) {
				result.succeeded++;
			} else {
				result.failed++;
				result.errors.push(`${deleteResult.recordId}: ${deleteResult.error}`);
			}
			completed++;
		}
		batchesCompleted++;

		// Calculate ETA based on average batch time
		let estimatedSecondsRemaining: number | undefined;
		if (batchTimes.length >= 2) {
			const avgBatchTime = batchTimes.reduce((a, b) => a + b, 0) / batchTimes.length;
			const remainingBatches = batches.length - batchesCompleted;
			estimatedSecondsRemaining = Math.ceil((avgBatchTime * remainingBatches) / 1000);
		}

		onProgress?.({
			completed,
			total,
			succeeded: result.succeeded,
			failed: result.failed,
			batchesCompleted,
			totalBatches: batches.length,
			estimatedSecondsRemaining,
		});
	}

	debugLog(
		"recordAPI",
		`✅ Delete complete: ${result.succeeded} succeeded, ${result.failed} failed`,
	);
	return result;
}

/**
 * Submit bulk delete for all records matching a FetchXML query
 * Use this when no specific records are selected - deletes ALL matching records
 *
 * @param fetchXml The FetchXML query defining which records to delete
 * @param jobName Name for the bulk delete job
 * @returns AsyncOperationId for tracking the job
 */
export async function submitBulkDeleteFromFetchXml(
	fetchXml: string,
	jobName: string,
): Promise<{ asyncOperationId: string; bulkDeleteOperationId: string; jobUrl: string }> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("recordAPI", `📡 Submitting bulk delete job from FetchXML: ${jobName}`);

	try {
		// Convert FetchXML to QueryExpression using the Dataverse function. BulkDelete only accepts
		// condition values as strings; the conversion returns numbers and booleans raw (#224)
		const queryExpression = withStringConditionValues(await fetchXmlToQueryExpression(fetchXml));

		// Submit BulkDelete with the converted QueryExpression
		const result = (await window.dataverseAPI!.execute({
			operationName: "BulkDelete",
			operationType: "action",
			parameters: {
				QuerySet: [queryExpression],
				JobName: jobName,
				SendEmailNotification: false,
				ToRecipients: [],
				CCRecipients: [],
				RecurrencePattern: "",
				StartDateTime: new Date().toISOString(),
			},
		})) as { JobId?: string; AsyncOperationId?: string };

		const asyncOpId = result.AsyncOperationId || result.JobId || "";

		// Query for the bulkdeleteoperationid using the asyncoperationid
		const bulkDeleteOpId = await getBulkDeleteOperationId(asyncOpId);

		// Build job tracking URL using bulkdeleteoperation entity
		const envUrl = await getEnvironmentUrl();
		const baseUrl = envUrl?.endsWith("/") ? envUrl.slice(0, -1) : envUrl;
		const jobUrl =
			baseUrl && bulkDeleteOpId
				? `${baseUrl}/main.aspx?pagetype=entityrecord&etn=bulkdeleteoperation&id=${bulkDeleteOpId}`
				: "";

		debugLog(
			"recordAPI",
			`✅ Bulk delete job submitted: asyncOpId=${asyncOpId}, bulkDeleteOpId=${bulkDeleteOpId}`,
		);

		return {
			asyncOperationId: asyncOpId,
			bulkDeleteOperationId: bulkDeleteOpId,
			jobUrl,
		};
	} catch (error) {
		console.error("submitBulkDeleteFromFetchXml: Failed:", error);
		throw error;
	}
}

// ============================================================================
// ACTIVATE / DEACTIVATE (#143)
// ============================================================================

interface StateOptionsResponse {
	OptionSet?: { Options?: Parameters<typeof toEntityStateModel>[0] };
}

/** Dataverse's answer when the entity has no such attribute (the statecode lookup) */
const isMissingAttributeError = (error: unknown) =>
	/0x80040217|could not find|does not exist|not found/i.test(error instanceof Error ? error.message : String(error));

/**
 * The entity's state model (#143): statecode options (StateAttributeMetadata, with DefaultStatus
 * and InvariantName), statuscode options (StatusAttributeMetadata, with State and TransitionData)
 * and EnforceStateTransitions. Null when the entity has no statecode. Memoized per connection.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query-metadata-web-api
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/define-custom-state-model-transitions
 */
export function getEntityStateModel(logicalName: string): Promise<EntityStateModel | null> {
	return memoizeForSession(`stateModel:${logicalName}`, async () => {
		if (!isDataverseAvailable()) throw new Error("PPTB Dataverse API not available");
		const entity = `EntityDefinitions(LogicalName='${odataString(logicalName)}')`;
		let states: StateOptionsResponse;
		try {
			states = (await window.dataverseAPI!.queryData(
				`${entity}/Attributes(LogicalName='statecode')/Microsoft.Dynamics.CRM.StateAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)`,
			)) as unknown as StateOptionsResponse;
		} catch (error) {
			if (isMissingAttributeError(error)) return null;
			throw error;
		}
		const [statuses, definition] = await Promise.all([
			window.dataverseAPI!.queryData(
				`${entity}/Attributes(LogicalName='statuscode')/Microsoft.Dynamics.CRM.StatusAttributeMetadata?$select=LogicalName&$expand=OptionSet($select=Options)`,
			) as unknown as Promise<StateOptionsResponse>,
			window.dataverseAPI!.queryData(`${entity}?$select=EnforceStateTransitions`) as unknown as Promise<{
				EnforceStateTransitions?: boolean;
			}>,
		]);
		const model = toEntityStateModel(
			states.OptionSet?.Options ?? [],
			statuses.OptionSet?.Options ?? [],
			definition.EnforceStateTransitions === true,
		);
		debugLog("metadataAPI", `State model for ${logicalName}:`, model);
		return model.states.length > 0 ? model : null;
	});
}

/**
 * Whether the user has the entity's Write privilege (#143). The privilege name comes from the
 * entity's metadata (Privileges, PrivilegeType "Write"): activities share prvWriteActivity, orders
 * use prvWriteOrder (verified on a sandbox). Undefined when it can't be determined: the action is
 * then allowed and Dataverse answers with its own error.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/retrieveuserprivilegebyprivilegename
 */
export function checkWritePrivilege(logicalName: string): Promise<boolean | undefined> {
	return checkEntityPrivilege(logicalName, "Write");
}

/**
 * Whether the user has the entity's privilege of this type (#143, #149), named from the entity's
 * metadata (Privileges, PrivilegeType). The close activities are activities, so Create is
 * prvCreateActivity (verified on a sandbox). Undefined when it can't be determined.
 */
export async function checkEntityPrivilege(logicalName: string, privilegeType: "Write" | "Create"): Promise<boolean | undefined> {
	if (!isDataverseAvailable()) return undefined;
	try {
		const user = await whoAmI();
		if (!user) return undefined;
		const definition = await memoizeForSession(
			`privileges:${logicalName}`,
			() =>
				window.dataverseAPI!.queryData(
					`EntityDefinitions(LogicalName='${odataString(logicalName)}')?$select=Privileges`,
				) as unknown as Promise<{ Privileges?: { Name?: string; PrivilegeType?: string }[] }>,
		);
		const name = definition.Privileges?.find((p) => p.PrivilegeType === privilegeType)?.Name;
		if (!name) return undefined;
		// Same session key as checkPrivilegeByName; a failed call isn't kept (#71)
		return await memoizeForSession(`privilege:${user.UserId}:${name}`, () => retrieveUserPrivilege(user.UserId, name));
	} catch (error) {
		console.warn(`checkEntityPrivilege(${logicalName}, ${privilegeType}): couldn't determine, allowing:`, error);
		return undefined;
	}
}

/**
 * The current statuscode of each record (#143), for status reason transitions. Undefined when it
 * can't be read. Filters by primary key in groups of 50 (short URLs).
 */
export async function getRecordStatuses(
	logicalName: string,
	primaryIdAttribute: string,
	recordIds: string[],
): Promise<number[] | undefined> {
	try {
		const entitySet = (await resolveEntitySet(logicalName))?.entitySetName;
		if (!entitySet) return undefined;
		const statuses: number[] = [];
		for (let i = 0; i < recordIds.length; i += 50) {
			const filter = recordIds
				.slice(i, i + 50)
				.map((id) => `${primaryIdAttribute} eq ${id.replace(/[{}]/g, "")}`)
				.join(" or ");
			const result = await window.dataverseAPI!.queryData(
				`${entitySet}?$select=statuscode&$filter=${encodeURIComponent(filter)}`,
			);
			for (const row of result.value) {
				if (typeof row.statuscode === "number") statuses.push(row.statuscode);
			}
		}
		return statuses;
	} catch (error) {
		console.warn(`getRecordStatuses(${logicalName}) failed:`, error);
		return undefined;
	}
}

/**
 * The current statecode of each record by ID (#149), in groups of 50 like getRecordStatuses.
 * Throws when it can't be read.
 */
export async function getRecordStateCodes(
	entitySetName: string,
	primaryIdAttribute: string,
	recordIds: string[],
): Promise<Map<string, number>> {
	const states = new Map<string, number>();
	const normalize = (id: string) => id.replace(/[{}]/g, "").toLowerCase();
	for (let i = 0; i < recordIds.length; i += 50) {
		const filter = recordIds
			.slice(i, i + 50)
			.map((id) => `${primaryIdAttribute} eq ${id.replace(/[{}]/g, "")}`)
			.join(" or ");
		const result = await window.dataverseAPI!.queryData(
			`${entitySetName}?$select=statecode&$filter=${encodeURIComponent(filter)}`,
		);
		for (const row of result.value) {
			const id = row[primaryIdAttribute];
			if (typeof id === "string" && typeof row.statecode === "number") states.set(normalize(id), row.statecode);
		}
	}
	// Keyed as the caller passed them
	const byCallerId = new Map<string, number>();
	for (const id of recordIds) {
		const state = states.get(normalize(id));
		if (state !== undefined) byCallerId.set(id, state);
	}
	return byCallerId;
}

/**
 * Activate / Deactivate (#143): sets `statecode` and `statuscode` on each record with
 * `PATCH {EntitySetName}(id)`. PPTB ≥ 1.2.7: $batch requests through executeBatch with the
 * metadata EntitySetName (like delete, #61), retried while throttled and stopped if the connection
 * changes (#66). Older hosts: one dataverseAPI.update per record, retried after service protection
 * errors (#96).
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/special-update-operation-behavior
 */
export async function setRecordsState(
	entityLogicalName: string,
	recordIds: string[],
	target: StateTarget,
	onProgress?: (progress: BatchDeleteProgress) => void,
): Promise<BatchDeleteResult> {
	if (!isDataverseAvailable()) throw new Error("PPTB Dataverse API not available");
	const shouldContinue = await connectionGuard();
	const entitySetName = await deleteEntitySet(entityLogicalName);
	debugLog(
		"recordAPI",
		`📡 Setting state ${target.statecode}/${target.statuscode} on ${recordIds.length} ${entityLogicalName} records`,
	);
	if (entitySetName) {
		return updateRecordStatesInBatches(entitySetName, recordIds, target, {
			executeBatch: hostExecuteBatch,
			onProgress,
			shouldContinue,
		});
	}

	return runOneByOne(
		recordIds,
		(recordId) => window.dataverseAPI!.update(entityLogicalName, recordId, statePatchBody(target)),
		STATE_CONNECTION_CHANGED_MESSAGE,
		shouldContinue,
		onProgress,
	);
}

/**
 * Hosts without executeBatch: one request per record, retried after service protection errors
 * (#96), stopped when the connection changes (#66), with the same progress shape as the batches.
 */
async function runOneByOne(
	recordIds: string[],
	send: (recordId: string) => Promise<unknown>,
	stoppedMessage: string,
	shouldContinue: () => boolean | Promise<boolean>,
	onProgress?: (progress: BatchDeleteProgress) => void,
): Promise<BatchDeleteResult> {
	const result: BatchDeleteResult = { succeeded: 0, failed: 0, errors: [] };
	const started = Date.now();
	for (const [index, recordId] of recordIds.entries()) {
		if (!(await shouldContinue())) {
			for (const id of recordIds.slice(index)) {
				result.failed++;
				result.errors.push(`${id}: ${stoppedMessage}`);
			}
			break;
		}
		try {
			// not retried after a connection change (#175)
			await withThrottleRetry(() => send(recordId), { shouldContinue });
			result.succeeded++;
		} catch (error) {
			result.failed++;
			result.errors.push(`${recordId}: ${formatDataverseError(error)}`);
		}
		const completed = index + 1;
		const average = (Date.now() - started) / completed;
		onProgress?.({
			completed,
			total: recordIds.length,
			succeeded: result.succeeded,
			failed: result.failed,
			batchesCompleted: completed,
			totalBatches: recordIds.length,
			estimatedSecondsRemaining:
				completed >= 2 ? Math.ceil((average * (recordIds.length - completed)) / 1000) : undefined,
		});
	}
	return result;
}

// ============================================================================
// CLOSE ACTIONS (#149)
// ============================================================================

export interface CompetitorOption {
	id: string;
	name: string;
	website?: string;
}

/** Competitors whose name contains the text (Lose opportunity), up to 25, by name */
export async function searchCompetitors(text: string): Promise<CompetitorOption[]> {
	if (!isDataverseAvailable()) return [];
	const term = text.trim();
	const filter = term ? `&$filter=${encodeURIComponent(`contains(name,'${odataString(term)}')`)}` : "";
	const result = await window.dataverseAPI!.queryData(
		`${COMPETITOR_ENTITY_SET}?$select=name,websiteurl&$orderby=name&$top=25${filter}`,
	);
	return result.value.flatMap((row) =>
		typeof row.competitorid === "string"
			? [{ id: row.competitorid, name: typeof row.name === "string" ? row.name : row.competitorid, website: typeof row.websiteurl === "string" ? row.websiteurl : undefined }]
			: [],
	);
}

/**
 * Runs a close message on each record (#149) with the same values: POST {action} (or the
 * collection-bound action) with the close activity bound to the record and Status. PPTB ≥ 1.2.7:
 * $batch requests through executeBatch with throttle retry and the connection guard; a batch that
 * stops early isn't redone for records it already closed, read from their statecode before and
 * after (completedByStoppedBatch). Older hosts: dataverseAPI.execute per record.
 * @see https://learn.microsoft.com/power-apps/developer/data-platform/webapi/use-web-api-actions
 */
export async function closeRecords(
	entityLogicalName: string,
	primaryIdAttribute: string | undefined,
	recordIds: string[],
	action: CloseMessageAction,
	values: CloseValues,
	onProgress?: (progress: BatchDeleteProgress) => void,
): Promise<BatchDeleteResult> {
	if (!isDataverseAvailable()) throw new Error("PPTB Dataverse API not available");
	const shouldContinue = await connectionGuard();
	const entitySetName = (await resolveEntitySet(entityLogicalName))?.entitySetName;
	if (!entitySetName) throw new Error(`The ${entityLogicalName} metadata has no entity set name.`);
	debugLog("recordAPI", `📡 ${action.message} on ${recordIds.length} ${entityLogicalName} records`);
	const parameters = (recordId: string) => buildCloseParameters(action, entitySetName, recordId, values);

	if (canExecuteBatch()) {
		let before: Map<string, number> | undefined;
		if (primaryIdAttribute) {
			try {
				before = await getRecordStateCodes(entitySetName, primaryIdAttribute, recordIds);
			} catch (error) {
				console.warn("Close: couldn't read the records' states first; a stopped batch is redone in full:", error);
			}
		}
		const url = closeRequestUrl(action, entitySetName);
		return closeRecordsInBatches(recordIds, (id) => ({ method: "POST", url, body: parameters(id) }), {
			executeBatch: hostExecuteBatch,
			onProgress,
			shouldContinue,
			completedInStoppedBatch:
				before && primaryIdAttribute
					? async (ids) =>
							completedByStoppedBatch(
								before!,
								await getRecordStateCodes(entitySetName, primaryIdAttribute, ids),
								action.targetState,
								ids,
							)
					: undefined,
		});
	}

	return runOneByOne(
		recordIds,
		(recordId) =>
			window.dataverseAPI!.execute({
				operationName: action.message,
				operationType: "action",
				// Collection-bound: the host builds {EntitySetName}/Microsoft.Dynamics.CRM.{operationName}
				entityName: action.boundToCollection ? entityLogicalName : undefined,
				parameters: parameters(recordId),
			}),
		CLOSE_CONNECTION_CHANGED_MESSAGE,
		shouldContinue,
		onProgress,
	);
}
