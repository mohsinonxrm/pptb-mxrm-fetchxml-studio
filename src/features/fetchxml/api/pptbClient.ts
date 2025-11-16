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
