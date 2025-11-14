/**
 * PPTB Dataverse API client wrapper
 * Provides typed interface to window.dataverseAPI exposed by PPTB host
 */

// Extend window interface for PPTB Dataverse API
declare global {
	interface Window {
		dataverseAPI?: {
			// CRUD operations
			create: (entitySetName: string, data: Record<string, unknown>) => Promise<{ id: string }>;
			retrieve: (
				entitySetName: string,
				id: string,
				select?: string[]
			) => Promise<Record<string, unknown>>;
			update: (entitySetName: string, id: string, data: Record<string, unknown>) => Promise<void>;
			delete: (entitySetName: string, id: string) => Promise<void>;

			// Query operations
			fetchXmlQuery: (fetchXml: string) => Promise<{
				value: Record<string, unknown>[];
				"@Microsoft.Dynamics.CRM.totalrecordcount"?: number;
				"@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"?: boolean;
				"@Microsoft.Dynamics.CRM.morerecords"?: boolean;
				"@Microsoft.Dynamics.CRM.pagingcookie"?: string;
			}>;
			retrieveMultiple: (
				entitySetName: string,
				query?: string
			) => Promise<{
				value: Record<string, unknown>[];
			}>;
			queryData: (odataQuery: string) => Promise<{
				value: Record<string, unknown>[];
			}>;

			// Metadata operations
			getEntityMetadata: (
				entityLogicalName: string,
				searchByLogicalName?: boolean,
				selectColumns?: string[]
			) => Promise<EntityMetadata>;
			getAllEntitiesMetadata: () => Promise<{ value: EntityMetadata[] }>;
			getEntityRelatedMetadata: (
				entityLogicalName: string,
				relatedPath: string,
				selectColumns?: string[]
			) => Promise<{ value: unknown[] }>;

			// Other operations
			execute: (options: {
				operationName: string;
				operationType: "action" | "function";
				entityName?: string;
				entityId?: string;
				parameters?: Record<string, unknown>;
			}) => Promise<unknown>;
			getSolutions: (selectColumns?: string[]) => Promise<{ value: unknown[] }>;

			// User context - note: WhoAmI is called via execute({ operationName: 'WhoAmI', operationType: 'function' })
		};
	}
}

import { debugLog } from "../../../shared/utils/debug";
import { responseHasFormattedValues } from "./formattedValues";

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
}

export interface AttributeMetadata {
	LogicalName: string;
	SchemaName: string;
	DisplayName?: { UserLocalizedLabel?: { Label?: string } };
	AttributeType: string;
	AttributeTypeName?: { Value?: string };
	MetadataId: string;
	IsValidForAdvancedFind?: { Value: boolean };
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
}

export interface FetchXmlResult {
	records: Record<string, unknown>[];
	totalRecordCount?: number;
	totalRecordCountLimitExceeded?: boolean;
	moreRecords?: boolean;
	pagingCookie?: string;
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

	return await window.dataverseAPI!.getEntityMetadata(logicalName, true);
}

/**
 * Get entity attributes metadata (filtered server-side for Advanced Find and sorted alphabetically)
 */
export async function getEntityAttributes(
	logicalName: string,
	advancedFindOnly: boolean = true
): Promise<AttributeMetadata[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// Build OData query with server-side filtering
	let query = `EntityDefinitions(LogicalName='${logicalName}')/Attributes?$select=LogicalName,SchemaName,DisplayName,AttributeType,AttributeTypeName,MetadataId,IsValidForAdvancedFind`;

	// Add server-side filter for IsValidForAdvancedFind
	if (advancedFindOnly) {
		query += "&$filter=IsValidForAdvancedFind/Value eq true";
	}

	debugLog("metadataAPI", `📡 GET Attributes for '${logicalName}': ${query}`);

	const result = await window.dataverseAPI!.queryData(query);

	const attributes = result.value as unknown as AttributeMetadata[];

	debugLog(
		"metadataAPI",
		`✅ Attributes retrieved for '${logicalName}': ${attributes.length} attributes`
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
 * Get a single attribute with full metadata including OptionSet expansion
 * Required for loading picklist/boolean option values
 * Uses type casting (BooleanAttributeMetadata or PicklistAttributeMetadata) to access OptionSet
 */
export async function getAttributeWithOptionSet(
	entityLogicalName: string,
	attributeLogicalName: string
): Promise<AttributeMetadata> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// First, get the basic attribute to determine its type
	const basicQuery = `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')?$select=AttributeType,LogicalName,SchemaName,DisplayName,MetadataId`;

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
	} else {
		// For non-optionset attributes, return the basic metadata
		debugLog(
			"metadataAPI",
			`⚠️ Attribute '${attributeLogicalName}' is type '${basicAttr.AttributeType}' - no OptionSet available`
		);
		return basicAttr;
	}

	// Build query with type cast and OptionSet expansion
	const fullQuery = `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')/${typeCast}?$expand=OptionSet`;

	debugLog(
		"metadataAPI",
		`📡 GET Attribute with OptionSet for '${attributeLogicalName}': ${fullQuery}`
	);

	const fullResult = await window.dataverseAPI!.queryData(fullQuery);
	// Single attribute query returns object directly, not in .value array
	const fullAttr = (fullResult.value?.[0] || fullResult) as unknown as AttributeMetadata;

	if (!fullAttr || !fullAttr.LogicalName) {
		throw new Error(
			`Failed to retrieve attribute ${attributeLogicalName} with OptionSet expansion`
		);
	}

	debugLog(
		"metadataAPI",
		`✅ Attribute with OptionSet retrieved for '${attributeLogicalName}':`,
		fullAttr.OptionSet
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
	attributeType: string
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
				`⚠️ Attribute '${attributeLogicalName}' is type '${attributeType}' - no detailed metadata needed`
			);
			// Return basic metadata query without type cast
			const basicQuery = `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')?$select=SchemaName,LogicalName,DisplayName,AttributeType,MetadataId`;
			const basicResult = await window.dataverseAPI!.queryData(basicQuery);
			return (basicResult.value?.[0] || basicResult) as unknown as AttributeMetadata;
		}
	}

	// Build query with type cast and specific property selection
	const detailedQuery = `EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')/${typeCast}?$select=${selectProps}`;

	debugLog(
		"metadataAPI",
		`📡 GET Detailed Attribute metadata for '${attributeLogicalName}' (${attributeType}): ${detailedQuery}`
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
		detailedAttr
	);

	return detailedAttr;
}

/**
 * Get entity relationships metadata (filtered server-side for Advanced Find and sorted alphabetically)
 */
export async function getEntityRelationships(
	logicalName: string,
	advancedFindOnly: boolean = true
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
		let query = `EntityDefinitions(LogicalName='${logicalName}')/${relType}?$select=SchemaName,RelationshipType,ReferencedEntity,ReferencedAttribute,ReferencingEntity,ReferencingAttribute,MetadataId,IsValidForAdvancedFind,IsCustomRelationship`;

		// Add server-side filter for IsValidForAdvancedFind
		if (advancedFindOnly) {
			query += "&$filter=IsValidForAdvancedFind eq true";
		}

		return query;
	};

	const buildManyToManyQuery = () => {
		let query = `EntityDefinitions(LogicalName='${logicalName}')/ManyToManyRelationships?$select=SchemaName,RelationshipType,Entity1LogicalName,Entity2LogicalName,Entity1IntersectAttribute,Entity2IntersectAttribute,IntersectEntityName,MetadataId,IsValidForAdvancedFind,IsCustomRelationship`;

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
		`✅ Relationships retrieved for '${logicalName}': 1:N=${oneToManyResult.value.length}, N:1=${manyToOneResult.value.length}, N:N=${manyToManyResult.value.length}`
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
 * Execute FetchXML query with formatted values
 * Formatted values are returned with @OData.Community.Display.V1.FormattedValue annotations
 */
export async function executeFetchXml(fetchXml: string): Promise<FetchXmlResult> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("fetchXmlAPI", `📡 Executing FetchXML query...`);
	debugLog("fetchXmlAPI", `FetchXML:\n${fetchXml}`);

	const result = await window.dataverseAPI!.fetchXmlQuery(fetchXml);

	// Log the raw response to console for inspection
	console.log("🔍 Raw FetchXML Response:", result);
	console.log("🔍 Sample Record (first):", result.value[0]);
	console.log("🔍 All Record Keys:", result.value[0] ? Object.keys(result.value[0]) : []);
	
	// Check if formatted values are included
	const hasFormatted = responseHasFormattedValues(result.value);
	console.log(`🔍 Response includes formatted values: ${hasFormatted}`);
	
	if (hasFormatted) {
		console.log("✅ Formatted values detected! Example formatted keys:");
		const formattedKeys = Object.keys(result.value[0] || {}).filter(k => k.includes("@OData"));
		console.log(formattedKeys.slice(0, 5));
	} else {
		console.warn("⚠️ NO FORMATTED VALUES DETECTED");
		console.warn("");
		console.warn("📋 The PPTB host's fetchXmlQuery() needs to update the Prefer header.");
		console.warn("");
		console.warn("🔧 In dataverseManager.ts (around line 405), change:");
		console.warn('   FROM: Prefer: "return=representation"');
		console.warn('   TO:   Prefer: "return=representation, odata.include-annotations=\\"OData.Community.Display.V1.FormattedValue\\""');
		console.warn("");
		console.warn("📖 Microsoft Learn reference:");
		console.warn("   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/fetchxml/select-columns?tabs=webapi#formatted-values");
		console.warn("");
		console.warn("💡 This will enable rich display: labels for picklists, names for lookups, formatted dates, etc.");
	}

	debugLog(
		"fetchXmlAPI",
		`✅ FetchXML query executed: ${result.value.length} records retrieved`
	);

	return {
		records: result.value,
		totalRecordCount: result["@Microsoft.Dynamics.CRM.totalrecordcount"],
		totalRecordCountLimitExceeded: result["@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"],
		moreRecords: result["@Microsoft.Dynamics.CRM.morerecords"],
		pagingCookie: result["@Microsoft.Dynamics.CRM.pagingcookie"],
	};
}

/**
 * Call WhoAmI to get current user context
 */
export async function whoAmI(): Promise<{
	UserId: string;
	BusinessUnitId: string;
	OrganizationId: string;
} | null> {
	if (!isDataverseAvailable()) {
		console.warn("PPTB Dataverse API not available");
		return null;
	}

	try {
		// Call WhoAmI using execute with proper signature from documentation
		const result = (await window.dataverseAPI!.execute({
			operationName: "WhoAmI",
			operationType: "function",
		})) as {
			UserId: string;
			BusinessUnitId: string;
			OrganizationId: string;
		};
		return result;
	} catch (error) {
		console.error("WhoAmI failed:", error);
		return null;
	}
}
