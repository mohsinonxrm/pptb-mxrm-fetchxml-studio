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
import { metadataCache } from "../state/cache";

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
		const formattedKeys = Object.keys(result.value[0] || {}).filter((k) => k.includes("@OData"));
		console.log(formattedKeys.slice(0, 5));
	} else {
		console.warn("⚠️ NO FORMATTED VALUES DETECTED");
		console.warn("");
		console.warn("📋 The PPTB host's fetchXmlQuery() needs to update the Prefer header.");
		console.warn("");
		console.warn("🔧 In dataverseManager.ts (around line 405), change:");
		console.warn('   FROM: Prefer: "return=representation"');
		console.warn(
			'   TO:   Prefer: "return=representation, odata.include-annotations=\\"OData.Community.Display.V1.FormattedValue\\""'
		);
		console.warn("");
		console.warn("📖 Microsoft Learn reference:");
		console.warn(
			"   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/fetchxml/select-columns?tabs=webapi#formatted-values"
		);
		console.warn("");
		console.warn(
			"💡 This will enable rich display: labels for picklists, names for lookups, formatted dates, etc."
		);
	}

	debugLog("fetchXmlAPI", `✅ FetchXML query executed: ${result.value.length} records retrieved`);

	return {
		records: result.value,
		totalRecordCount: result["@Microsoft.Dynamics.CRM.totalrecordcount"],
		totalRecordCountLimitExceeded: result["@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded"],
		moreRecords: result["@Microsoft.Dynamics.CRM.morerecords"],
		pagingCookie: result["@Microsoft.Dynamics.CRM.pagingcookie"],
	};
}

/**
 * Execute a System View (savedquery) by its ID
 * Uses the optimized ?savedQuery={id} parameter instead of fetching and executing fetchXml
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/saved-queries#use-a-saved-query-in-fetchxml
 */
export async function executeSystemView(
	entitySetName: string,
	savedQueryId: string
): Promise<FetchXmlResult> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("fetchXmlAPI", `📡 Executing System View: ${savedQueryId} on ${entitySetName}`);

	// Use queryData with savedQuery parameter (OData query, not FetchXML)
	// GET [Organization URI]/api/data/v9.2/{entitySetName}?savedQuery={savedQueryId}
	const query = `${entitySetName}?savedQuery=${savedQueryId}`;
	const result = await window.dataverseAPI!.queryData(query);

	debugLog("fetchXmlAPI", `✅ System View executed: ${result.value.length} records retrieved`);

	return {
		records: result.value,
		// savedQuery execution doesn't return paging metadata in the same way
		totalRecordCount: undefined,
		moreRecords: undefined,
		pagingCookie: undefined,
	};
}

/**
 * Execute a Personal View (userquery) by its ID
 * Uses the optimized ?userQuery={id} parameter instead of fetching and executing fetchXml
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/saved-queries#use-a-saved-query-in-fetchxml
 */
export async function executePersonalView(
	entitySetName: string,
	userQueryId: string
): Promise<FetchXmlResult> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("fetchXmlAPI", `📡 Executing Personal View: ${userQueryId} on ${entitySetName}`);

	// Use queryData with userQuery parameter (OData query, not FetchXML)
	// GET [Organization URI]/api/data/v9.2/{entitySetName}?userQuery={userQueryId}
	const query = `${entitySetName}?userQuery=${userQueryId}`;
	const result = await window.dataverseAPI!.queryData(query);

	debugLog("fetchXmlAPI", `✅ Personal View executed: ${result.value.length} records retrieved`);

	return {
		records: result.value,
		// userQuery execution doesn't return paging metadata in the same way
		totalRecordCount: undefined,
		moreRecords: undefined,
		pagingCookie: undefined,
	};
}

/**
 * Call WhoAmI to get current user context
 */
export async function whoAmI(): Promise<WhoAmIResponse | null> {
	if (!isDataverseAvailable()) {
		console.warn("PPTB Dataverse API not available");
		return null;
	}

	try {
		// Call WhoAmI using execute with proper signature from documentation
		const result = (await window.dataverseAPI!.execute({
			operationName: "WhoAmI",
			operationType: "function",
		})) as WhoAmIResponse;
		return result;
	} catch (error) {
		console.error("WhoAmI failed:", error);
		return null;
	}
}

/**
 * Check if user has a specific privilege using RetrieveUserPrivilegeByPrivilegeName
 */
export async function checkPrivilegeByName(
	userId: string,
	privilegeName: string
): Promise<boolean> {
	if (!isDataverseAvailable()) {
		return false;
	}

	try {
		const query = `systemusers(${userId})/Microsoft.Dynamics.CRM.RetrieveUserPrivilegeByPrivilegeName(PrivilegeName='${privilegeName}')`;
		const result = await window.dataverseAPI!.queryData(query);

		// Log raw response for debugging
		console.log(`🔍 checkPrivilegeByName(${privilegeName}) raw response:`, result);

		// RetrieveUserPrivilegeByPrivilegeName returns the response directly, not wrapped in .value
		const response = result as unknown as PrivilegeCheckResponse;

		const hasPrivilege = !!response?.RolePrivileges?.length;
		console.log(
			`✅ checkPrivilegeByName(${privilegeName}): ${hasPrivilege ? "GRANTED" : "DENIED"}`
		);

		return hasPrivilege;
	} catch (error) {
		console.error(`checkPrivilegeByName(${privilegeName}) failed:`, error);
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
				0
			)} total solutions`
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
			})
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
		const query = `solutions?${SELECT}&$filter=${IS_VISIBLE} and ${HAS_ENTITIES}&$orderby=friendlyname asc`;

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

		console.log("[API] getSolutionComponents - Querying solutions individually:", solutionIds);

		// Query each solution individually (virtual entity doesn't support OR filters properly)
		const allComponents = await Promise.all(
			solutionIds.map(async (solutionId) => {
				const query = `msdyn_solutioncomponentsummaries?${SELECT}&$filter=msdyn_componenttype eq ${COMPONENT_TYPE_ENTITY} and msdyn_solutionid eq '${solutionId}'`;

				console.log("[API] getSolutionComponents - Query for solution:", { solutionId, query });
				debugLog("solutionComponentAPI", `📡 GET Solution components for solution: ${solutionId}`);
				const result = await window.dataverseAPI!.queryData(query);
				console.log("[API] getSolutionComponents - Result for solution:", {
					solutionId,
					componentCount: result.value.length,
					components: result.value,
				});
				return result.value as unknown as SolutionComponent[];
			})
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

		console.log("[API] getSolutionComponents - Combined results:", {
			totalSolutions: solutionIds.length,
			uniqueComponents: components.length,
			componentNames: components.map((c) => c.msdyn_name),
		});

		debugLog(
			"solutionComponentAPI",
			`✅ Components retrieved: ${components.length} unique entities`
		);
		return components;
	} catch (error) {
		console.error("getSolutionComponents: API call failed:", error);
		throw error;
	}
}

/**
 * Get ALL EntityDefinitions that are valid for Advanced Find
 * This is called once on startup and cached globally for the session
 */
export async function getAllAdvancedFindEntities(): Promise<EntityMetadata[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	// Check cache first
	const cached = metadataCache.getAllEntityMetadata();
	if (cached) {
		console.log("[API] getAllAdvancedFindEntities - Cache hit:", cached.length, "entities");
		return cached;
	}

	// Check for in-flight promise
	const inFlight = metadataCache.getAllEntityMetadataPromise();
	if (inFlight) {
		console.log("[API] getAllAdvancedFindEntities - Returning in-flight promise");
		return inFlight;
	}

	console.log("[API] getAllAdvancedFindEntities - Fetching all AF-valid entities");

	try {
		const promise = getAllEntities(true);
		metadataCache.setAllEntityMetadataPromise(promise);

		const entities = await promise;
		console.log("[API] getAllAdvancedFindEntities - Fetched:", entities.length, "entities");

		metadataCache.setAllEntityMetadata(entities);
		metadataCache.clearAllEntityMetadataPromise();

		return entities;
	} catch (error) {
		metadataCache.clearAllEntityMetadataPromise();
		console.error("[API] getAllAdvancedFindEntities - Failed:", error);
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

	console.log("[API] filterCachedEntitiesByNames - Filtered:", {
		requestedCount: logicalNames.length,
		requestedNames: logicalNames,
		filteredCount: filtered.length,
		filteredNames: filtered.map((e: EntityMetadata) => e.LogicalName),
	});

	return filtered;
}

/**
 * Get EntityDefinitions filtered by logical names and IsValidForAdvancedFind
 */
export async function getAdvancedFindEntitiesByNames(
	logicalNames: string[]
): Promise<EntityMetadata[]> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	if (!logicalNames.length) {
		console.log(
			"[API] getAdvancedFindEntitiesByNames - Empty logical names, returning all AF entities"
		);
		// Return all AF-valid entities
		return getAllEntities(true);
	}

	console.log("[API] getAdvancedFindEntitiesByNames - Input:", {
		count: logicalNames.length,
		names: logicalNames,
	});

	try {
		const SELECT =
			"$select=LogicalName,SchemaName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,IsValidForAdvancedFind,MetadataId,ObjectTypeCode";

		// Chunk logical names to avoid URL length limits
		const chunkSize = 15;
		const chunks: string[][] = [];
		for (let i = 0; i < logicalNames.length; i += chunkSize) {
			chunks.push(logicalNames.slice(i, i + chunkSize));
		}

		const allEntities = await Promise.all(
			chunks.map(async (chunk) => {
				const nameFilter = chunk.map((name) => `LogicalName eq '${name}'`).join(" or ");
				const query = `EntityDefinitions?${SELECT}&$filter=IsValidForAdvancedFind eq true and (${nameFilter})`;

				console.log("[API] getAdvancedFindEntitiesByNames - Query:", query);
				debugLog("metadataAPI", `📡 GET Entities by names: ${chunk.length} names`);
				const result = await window.dataverseAPI!.queryData(query);
				const entities = result.value as unknown as EntityMetadata[];
				console.log("[API] getAdvancedFindEntitiesByNames - Result:", {
					requestedNames: chunk,
					returnedCount: entities.length,
					returnedEntities: entities.map((e) => e.LogicalName),
				});
				return entities;
			})
		);

		// Flatten and dedupe
		const entityMap = new Map<string, EntityMetadata>();
		allEntities.flat().forEach((e) => entityMap.set(e.LogicalName, e));
		const entities = Array.from(entityMap.values());

		// Sort by display name
		entities.sort((a, b) => {
			const nameA = a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName;
			const nameB = b.DisplayName?.UserLocalizedLabel?.Label || b.LogicalName;
			return nameA.localeCompare(nameB);
		});

		debugLog("metadataAPI", `✅ Entities retrieved: ${entities.length}`);
		return entities;
	} catch (error) {
		console.error("getAdvancedFindEntitiesByNames: API call failed:", error);
		throw error;
	}
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
	/** Original FetchXML from the view - used for comparison */
	originalFetchXml: string;
	/** Entity set name for the execution URL */
	entitySetName: string;
	/** View name for display purposes */
	name: string;
	/** Optional LayoutXML for column configuration */
	layoutxml?: string;
}

/**
 * Parsed layout column from LayoutXML
 */
export interface LayoutColumn {
	name: string;
	width: number;
	disableSorting?: boolean;
	/** For link-entity columns, the alias prefix (e.g., "accountprimarycontactidcontactcontactid") */
	linkEntityAlias?: string;
	/** The actual attribute name without alias prefix */
	attributeName: string;
}

/**
 * Parsed LayoutXML structure
 */
export interface ParsedLayoutXml {
	gridName: string;
	objectTypeCode: number;
	jumpAttribute?: string;
	primaryIdAttribute: string;
	columns: LayoutColumn[];
}

/**
 * Parse LayoutXML into a structured format for DataGrid column configuration
 * @param layoutXml The raw LayoutXML string from savedquery/userquery
 * @returns Parsed layout structure or null if parsing fails
 */
export function parseLayoutXml(layoutXml: string): ParsedLayoutXml | null {
	if (!layoutXml) return null;

	try {
		const parser = new DOMParser();
		const doc = parser.parseFromString(layoutXml, "text/xml");

		const gridElement = doc.querySelector("grid");
		if (!gridElement) return null;

		const rowElement = doc.querySelector("row");
		if (!rowElement) return null;

		const columns: LayoutColumn[] = [];
		const cellElements = doc.querySelectorAll("row > cell");

		cellElements.forEach((cell) => {
			const name = cell.getAttribute("name") || "";
			const width = parseInt(cell.getAttribute("width") || "100", 10);
			const disableSorting = cell.getAttribute("disableSorting") === "1";

			// Check if this is a link-entity column (contains a dot)
			let linkEntityAlias: string | undefined;
			let attributeName = name;

			if (name.includes(".")) {
				const parts = name.split(".");
				linkEntityAlias = parts[0];
				attributeName = parts[1];
			}

			columns.push({
				name,
				width,
				disableSorting: disableSorting || undefined,
				linkEntityAlias,
				attributeName,
			});
		});

		return {
			gridName: gridElement.getAttribute("name") || "resultset",
			objectTypeCode: parseInt(gridElement.getAttribute("object") || "0", 10),
			jumpAttribute: gridElement.getAttribute("jump") || undefined,
			primaryIdAttribute: rowElement.getAttribute("id") || "",
			columns,
		};
	} catch (error) {
		console.error("parseLayoutXml: Failed to parse LayoutXML:", error);
		return null;
	}
}

/**
 * Generate LayoutXML from column configuration
 * Used when saving a customized view
 * @param columns The column configuration
 * @param objectTypeCode The entity's object type code
 * @param primaryIdAttribute The entity's primary ID attribute (e.g., "accountid")
 * @param jumpAttribute Optional attribute for the "jump" field
 */
export function generateLayoutXml(
	columns: LayoutColumn[],
	objectTypeCode: number,
	primaryIdAttribute: string,
	jumpAttribute?: string
): string {
	const cellsXml = columns
		.map((col) => {
			let cellAttrs = `name="${col.name}" width="${col.width}"`;
			if (col.disableSorting) {
				cellAttrs += ` disableSorting="1"`;
			}
			return `<cell ${cellAttrs} />`;
		})
		.join("");

	const jumpAttr = jumpAttribute ? ` jump="${jumpAttribute}"` : "";

	return (
		`<grid name="resultset" object="${objectTypeCode}"${jumpAttr} select="1" icon="1" preview="1">` +
		`<row name="result" id="${primaryIdAttribute}">${cellsXml}</row></grid>`
	);
}

/**
 * Get System Views (savedquery) for an entity
 * Only returns public views (querytype = 0) that are active and not hidden
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
		// querytype = 0: Public views (not lookup views, quickfind, etc.)
		// componentstate = 0: Published (not unpublished/deleted)
		// statecode = 0: Active views only
		const query =
			`savedqueries?$select=savedqueryid,name,fetchxml,layoutxml,isdefault` +
			`&$filter=(returnedtypecode eq '${entityLogicalName}' and querytype eq 0 and componentstate eq 0 and statecode eq 0)` +
			`&$orderby=isdefault desc,name asc`;

		debugLog("viewAPI", `Query: ${query}`);

		const result = await window.dataverseAPI!.queryData(query);

		const views: SavedView[] = (result.value || []).map((record) => ({
			id: record.savedqueryid as string,
			name: record.name as string,
			fetchxml: record.fetchxml as string,
			layoutxml: record.layoutxml as string | undefined,
			type: "system" as const,
			isDefault: record.isdefault as boolean | undefined,
		}));

		debugLog(
			"viewAPI",
			`✅ System Views retrieved for '${entityLogicalName}': ${views.length} views`
		);

		return views;
	} catch (error) {
		console.error(`getSystemViews: Failed to get system views for '${entityLogicalName}':`, error);
		throw error;
	}
}

/**
 * Get Personal Views (userquery) for an entity
 * Returns views owned by the current user or their teams
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
		// querytype = 0: Saved views (not quick find)
		// EqualUserOrUserTeams: Views owned by current user or their teams
		const query =
			`userqueries?$select=userqueryid,name,fetchxml,layoutxml` +
			`&$filter=(statecode eq 0 and querytype eq 0 and returnedtypecode eq '${entityLogicalName}' and Microsoft.Dynamics.CRM.EqualUserOrUserTeams(PropertyName='ownerid'))` +
			`&$orderby=name asc`;

		debugLog("viewAPI", `Query: ${query}`);

		const result = await window.dataverseAPI!.queryData(query);

		const views: SavedView[] = (result.value || []).map((record) => ({
			id: record.userqueryid as string,
			name: record.name as string,
			fetchxml: record.fetchxml as string,
			layoutxml: record.layoutxml as string | undefined,
			type: "personal" as const,
		}));

		debugLog(
			"viewAPI",
			`✅ Personal Views retrieved for '${entityLogicalName}': ${views.length} views`
		);

		return views;
	} catch (error) {
		console.error(
			`getPersonalViews: Failed to get personal views for '${entityLogicalName}':`,
			error
		);
		throw error;
	}
}

/**
 * Get all views (System + Personal) for an entity
 * @param entityLogicalName The logical name of the entity
 */
export async function getAllViews(
	entityLogicalName: string
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
	fetchXml: string
): Promise<ValidateFetchXmlExpressionResponse> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 ValidateFetchXmlExpression - Validating FetchXML...`);

	try {
		// The function requires URL-encoded FetchXML passed as a query parameter
		// Format: ValidateFetchXmlExpression(FetchXml=@FetchXml)?@FetchXml='encodedFetchXml'
		const encodedFetchXml = encodeURIComponent(fetchXml);
		const query = `ValidateFetchXmlExpression(FetchXml=@FetchXml)?@FetchXml='${encodedFetchXml}'`;

		const result = (await window.dataverseAPI!.queryData(
			query
		)) as unknown as ValidateFetchXmlExpressionResponse;

		debugLog(
			"viewAPI",
			`✅ ValidateFetchXmlExpression - Complete: ${
				result.ValidationResults?.Messages?.length || 0
			} messages`
		);

		return result;
	} catch (error) {
		console.error("validateFetchXmlExpression: Failed:", error);
		throw error;
	}
}

/**
 * Validate SavedQuery before creation/update
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/validatesavedquery
 */
export async function validateSavedQuery(fetchXml: string, queryType: number = 0): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 ValidateSavedQuery - Validating saved query...`);

	try {
		await window.dataverseAPI!.execute({
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
		`✅ SavedQuery privileges: writeQuery=${canWriteQuery}, writeCustomization=${canWriteCustomization}, publishCustomization=${canPublishCustomization} → canWrite=${canWrite}, canPublish=${canPublish}`
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
	description?: string;
	querytype?: number; // 0 = Main/Public view
}

/**
 * Data for creating a new personal view (userquery)
 */
export interface CreateUserQueryData {
	name: string;
	fetchxml: string;
	layoutxml: string;
	returnedtypecode: string; // Entity logical name
	description?: string;
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
	data: Partial<CreateSavedQueryData>
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
	data: Partial<CreateUserQueryData>
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
	addRequiredComponents: boolean = false
): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog(
		"viewAPI",
		`📡 AddSolutionComponent: ${componentId} (type=${componentType}) to "${solutionUniqueName}"`
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
 * Publish a specific savedquery
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/publishxml
 */
export async function publishSavedQuery(savedQueryId: string): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 PublishXml for SavedQuery: ${savedQueryId}`);

	try {
		// Build ParameterXml to publish a specific savedquery
		const parameterXml = `<importexportxml><savedqueries><savedquery>{${savedQueryId}}</savedquery></savedqueries></importexportxml>`;

		await window.dataverseAPI!.execute({
			operationName: "PublishXml",
			operationType: "action",
			parameters: {
				ParameterXml: parameterXml,
			},
		});

		debugLog("viewAPI", `✅ SavedQuery published: ${savedQueryId}`);
	} catch (error) {
		console.error("publishSavedQuery: Failed:", error);
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
		const query = `solutions?$select=solutionid,ismanaged&$filter=uniquename eq '${solutionUniqueName}'`;
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
	solutionUniqueName: string
): Promise<string | null> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("viewAPI", `📡 Get solution ID for "${solutionUniqueName}"`);

	try {
		const query = `solutions?$select=solutionid&$filter=uniquename eq '${solutionUniqueName}'`;
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
	viewName: string
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
		`📡 FetchXML length: ${fetchXml.length}, LayoutXML length: ${layoutXml.length}`
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
			`✅ ExportToExcel successful, file size: ${result.ExcelFile?.length || 0} bytes`
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
export function downloadBase64File(
	base64Data: string,
	filename: string,
	mimeType: string = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
): void {
	try {
		// Convert Base64 to Blob
		const byteCharacters = atob(base64Data);
		const byteNumbers = new Array(byteCharacters.length);
		for (let i = 0; i < byteCharacters.length; i++) {
			byteNumbers[i] = byteCharacters.charCodeAt(i);
		}
		const byteArray = new Uint8Array(byteNumbers);
		const blob = new Blob([byteArray], { type: mimeType });

		// Create download link
		const url = URL.createObjectURL(blob);
		const link = document.createElement("a");
		link.href = url;
		link.download = filename;

		// Trigger download
		document.body.appendChild(link);
		link.click();
		document.body.removeChild(link);

		// Clean up
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
		if (typeof window !== "undefined" && window.toolboxAPI?.connections?.getActiveConnection) {
			const conn = await window.toolboxAPI.connections.getActiveConnection();
			return conn?.url || null;
		}
		return null;
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
	environmentUrl: string
): string {
	// Remove trailing slash from environmentUrl if present to avoid double slashes
	const baseUrl = environmentUrl.endsWith("/") ? environmentUrl.slice(0, -1) : environmentUrl;
	// Standard Dynamics 365 record URL format
	return `${baseUrl}/main.aspx?etn=${entityName}&id=${recordId}&pagetype=entityrecord`;
}

/**
 * Delete a single record
 * @param entitySetName OData entity set name (e.g., "accounts")
 * @param recordId GUID of the record to delete
 */
export async function deleteRecord(entitySetName: string, recordId: string): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("recordAPI", `📡 Deleting record: ${entitySetName}(${recordId})`);

	try {
		await window.dataverseAPI!.delete(entitySetName, recordId);
		debugLog("recordAPI", `✅ Record deleted: ${entitySetName}(${recordId})`);
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
 * Convert FetchXML to QueryExpression using the FetchXmlToQueryExpression function
 * This function must be called via GET request with URL-encoded parameters
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/fetchxmltoqueryexpression
 */
export async function fetchXmlToQueryExpression(fetchXml: string): Promise<unknown> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("recordAPI", `📡 Converting FetchXML to QueryExpression`);

	try {
		// FetchXmlToQueryExpression is a function - must be called via GET with URL-encoded FetchXml
		// The parameter must be properly escaped and wrapped in single quotes
		const encodedFetchXml = encodeURIComponent(fetchXml);
		const functionUrl = `FetchXmlToQueryExpression(FetchXml=@p1)?@p1='${encodedFetchXml}'`;

		const result = await window.dataverseAPI!.queryData(functionUrl);

		debugLog("recordAPI", `✅ Converted FetchXML to QueryExpression`, result);
		return (result as { Query?: unknown }).Query || result;
	} catch (error) {
		console.error("fetchXmlToQueryExpression: Failed:", error);
		throw error;
	}
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
				`✅ Found bulkdeleteoperationid: ${operations[0].bulkdeleteoperationid}`
			);
			return operations[0].bulkdeleteoperationid;
		}

		debugLog(
			"recordAPI",
			`⚠️ No bulkdeleteoperationid found for asyncoperationid: ${asyncOperationId}`
		);
		return "";
	} catch (error) {
		console.error("getBulkDeleteOperationId: Failed:", error);
		return "";
	}
}

/**
 * Submit a bulk delete job for selected records
 * Uses FetchXmlToQueryExpression function to convert FetchXML to QueryExpression
 * @see https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/fetchxmltoqueryexpression
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
	jobName: string
): Promise<{ asyncOperationId: string; bulkDeleteOperationId: string; jobUrl: string }> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog(
		"recordAPI",
		`📡 Submitting bulk delete job: ${jobName} for ${recordIds.length} records`
	);

	try {
		// Build FetchXML to select records by their IDs
		// Use <value> elements for the "in" operator
		const valuesXml = recordIds.map((id) => `<value>${id}</value>`).join("");
		const fetchXml = `<fetch>
			<entity name="${entityLogicalName}">
				<attribute name="${primaryIdAttribute}" />
				<filter>
					<condition attribute="${primaryIdAttribute}" operator="in">
						${valuesXml}
					</condition>
				</filter>
			</entity>
		</fetch>`;

		// Convert FetchXML to QueryExpression using the Dataverse function
		const queryExpression = await fetchXmlToQueryExpression(fetchXml);

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
			`✅ Bulk delete job submitted: asyncOpId=${asyncOpId}, bulkDeleteOpId=${bulkDeleteOpId}`
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
		const odataQuery = `workflows?$select=workflowid,name,description,primaryentity,category,type,statuscode&$filter=(statecode eq 1 and statuscode eq 2 and ondemand eq true and category eq 0 and type eq 1 and primaryentity eq '${entityLogicalName}')&$orderby=name asc`;

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
	entityLogicalName: string
): Promise<void> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog(
		"workflowAPI",
		`📡 Executing workflow ${workflowId} on ${entityLogicalName}(${recordId})`
	);

	try {
		await window.dataverseAPI!.execute({
			operationName: "ExecuteWorkflow",
			operationType: "action",
			entityName: "workflow",
			entityId: workflowId,
			parameters: {
				EntityId: recordId,
			},
		});

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
	batchSize: number = 10
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
		`📡 Executing workflow ${workflowId} on ${recordIds.length} records in ${batches.length} batches`
	);

	// Track timing for ETA calculation
	const batchTimes: number[] = [];
	let batchesCompleted = 0;

	// Process batches sequentially (each batch processes in parallel)
	for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
		const batchRecordIds = batches[batchIndex];
		const batchStartTime = Date.now();

		debugLog(
			"workflowAPI",
			`📡 Executing batch ${batchIndex + 1}/${batches.length} with ${batchRecordIds.length} records`
		);

		// Execute workflows in parallel within the batch
		const workflowPromises = batchRecordIds.map(async (recordId) => {
			try {
				await executeWorkflow(workflowId, recordId, entityLogicalName);
				return { success: true, recordId };
			} catch (error) {
				return {
					success: false,
					recordId,
					error: error instanceof Error ? error.message : String(error),
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
		`✅ Workflow execution complete: ${result.succeeded} succeeded, ${result.failed} failed`
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
 * Delete multiple records using parallel delete requests via PPTB API
 * Recommended for 1-100 records. For larger sets, use submitBulkDelete.
 * Uses batches of parallel requests for performance while respecting rate limits.
 *
 * @param entitySetName OData entity set name (e.g., "accounts")
 * @param recordIds Array of record GUIDs to delete
 * @param onProgress Callback for progress updates with ETA
 * @param batchSize Number of concurrent requests per batch (default: 10)
 */
export async function deleteRecordsBatch(
	entitySetName: string,
	recordIds: string[],
	onProgress?: (progress: BatchDeleteProgress) => void,
	batchSize: number = 10
): Promise<BatchDeleteResult> {
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

	// Process batches sequentially (each batch processes in parallel)
	for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
		const batchRecordIds = batches[batchIndex];
		const batchStartTime = Date.now();

		// Execute deletes in parallel within the batch
		const deletePromises = batchRecordIds.map(async (recordId) => {
			try {
				await deleteRecord(entitySetName, recordId);
				return { success: true, recordId };
			} catch (error) {
				return {
					success: false,
					recordId,
					error: error instanceof Error ? error.message : String(error),
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
		`✅ Delete complete: ${result.succeeded} succeeded, ${result.failed} failed`
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
	jobName: string
): Promise<{ asyncOperationId: string; bulkDeleteOperationId: string; jobUrl: string }> {
	if (!isDataverseAvailable()) {
		throw new Error("PPTB Dataverse API not available");
	}

	debugLog("recordAPI", `📡 Submitting bulk delete job from FetchXML: ${jobName}`);

	try {
		// Convert FetchXML to QueryExpression using the Dataverse function
		const queryExpression = await fetchXmlToQueryExpression(fetchXml);

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
			`✅ Bulk delete job submitted: asyncOpId=${asyncOpId}, bulkDeleteOpId=${bulkDeleteOpId}`
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
