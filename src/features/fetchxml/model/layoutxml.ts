/**
 * LayoutXML Parser and Generator
 * Handles Dataverse view layout configuration
 *
 * LayoutXML structure:
 * <grid name="resultset" object="1" jump="name" select="1" icon="1" preview="1">
 *   <row name="result" id="accountid">
 *     <cell name="name" width="300" />
 *     <cell name="primarycontactid" width="150" />
 *     <cell name="createdon" width="125" />
 *   </row>
 * </grid>
 */

import type { FetchNode, AttributeNode, LinkEntityNode } from "./nodes";

/**
 * Column configuration for a single column in the grid
 */
export interface LayoutColumn {
	/** Attribute logical name */
	name: string;
	/** Column width in pixels */
	width: number;
	/** Link-entity alias if this column is from a related entity */
	linkEntityAlias?: string;
	/** Whether sorting is disabled for this column */
	disableSorting?: boolean;
	/** Image provider for icon columns */
	imageProviderName?: string;
	/** Original display name from metadata (for UI) */
	displayName?: string;
}

/**
 * Full layout configuration for the grid
 */
export interface LayoutXmlConfig {
	/** Grid name (usually "resultset") */
	gridName: string;
	/** Entity type code */
	objectTypeCode?: number;
	/** Attribute to navigate to on double-click */
	jumpAttribute?: string;
	/** Row selection enabled */
	enableSelection?: boolean;
	/** Show entity icon */
	showIcon?: boolean;
	/** Enable row preview */
	enablePreview?: boolean;
	/** Primary ID attribute for the row */
	primaryIdAttribute?: string;
	/** Column configurations in display order */
	columns: LayoutColumn[];
}

/**
 * Default column widths based on attribute type
 */
const DEFAULT_WIDTHS: Record<string, number> = {
	// Text types - wider
	String: 200,
	Memo: 250,

	// Numeric types - narrower
	Integer: 100,
	BigInt: 120,
	Decimal: 120,
	Double: 120,
	Money: 130,

	// Date/Time
	DateTime: 150,

	// Lookups - medium
	Lookup: 180,
	Customer: 180,
	Owner: 180,

	// Option sets - medium
	Picklist: 150,
	State: 100,
	Status: 120,
	MultiSelectPicklist: 200,

	// Boolean - narrow
	Boolean: 80,

	// Unique identifier - wide (GUIDs)
	Uniqueidentifier: 280,

	// Image - narrow
	Image: 60,

	// File - medium
	File: 150,

	// Default fallback
	default: 150,
};

/**
 * Get default width for an attribute type
 */
export function getDefaultWidthForType(attributeType?: string): number {
	if (!attributeType) return DEFAULT_WIDTHS.default;
	return DEFAULT_WIDTHS[attributeType] ?? DEFAULT_WIDTHS.default;
}

/**
 * Parse a LayoutXML string into a LayoutXmlConfig
 */
export function parseLayoutXml(xml: string): LayoutXmlConfig {
	const parser = new DOMParser();
	const doc = parser.parseFromString(xml, "text/xml");

	// Check for parsing errors
	const parseError = doc.querySelector("parsererror");
	if (parseError) {
		throw new Error(`Invalid LayoutXML: ${parseError.textContent}`);
	}

	const gridElement = doc.querySelector("grid");
	if (!gridElement) {
		throw new Error("Invalid LayoutXML: missing <grid> element");
	}

	const rowElement = doc.querySelector("row");
	const cellElements = doc.querySelectorAll("cell");

	const columns: LayoutColumn[] = [];
	cellElements.forEach((cell) => {
		const name = cell.getAttribute("name");
		const width = cell.getAttribute("width");
		const disableSorting = cell.getAttribute("disableSorting");
		const imageProviderName = cell.getAttribute("imageproviderwebresource");

		if (name) {
			columns.push({
				name,
				width: width ? parseInt(width, 10) : DEFAULT_WIDTHS.default,
				disableSorting: disableSorting === "1",
				imageProviderName: imageProviderName || undefined,
			});
		}
	});

	return {
		gridName: gridElement.getAttribute("name") || "resultset",
		objectTypeCode: gridElement.getAttribute("object")
			? parseInt(gridElement.getAttribute("object")!, 10)
			: undefined,
		jumpAttribute: gridElement.getAttribute("jump") || undefined,
		enableSelection: gridElement.getAttribute("select") === "1",
		showIcon: gridElement.getAttribute("icon") === "1",
		enablePreview: gridElement.getAttribute("preview") === "1",
		primaryIdAttribute: rowElement?.getAttribute("id") || undefined,
		columns,
	};
}

/**
 * Generate LayoutXML string from a LayoutXmlConfig
 */
export function generateLayoutXml(config: LayoutXmlConfig): string {
	const cells = config.columns
		.map((col) => {
			const attrs: string[] = [`name="${col.name}"`, `width="${col.width}"`];
			if (col.disableSorting) {
				attrs.push('disableSorting="1"');
			}
			if (col.imageProviderName) {
				attrs.push(`imageproviderwebresource="${col.imageProviderName}"`);
			}
			return `    <cell ${attrs.join(" ")} />`;
		})
		.join("\n");

	const gridAttrs: string[] = [`name="${config.gridName}"`];
	if (config.objectTypeCode !== undefined) {
		gridAttrs.push(`object="${config.objectTypeCode}"`);
	}
	if (config.jumpAttribute) {
		gridAttrs.push(`jump="${config.jumpAttribute}"`);
	}
	gridAttrs.push(`select="${config.enableSelection ? "1" : "0"}"`);
	gridAttrs.push(`icon="${config.showIcon ? "1" : "0"}"`);
	gridAttrs.push(`preview="${config.enablePreview ? "1" : "0"}"`);

	const rowId = config.primaryIdAttribute || "id";

	return `<grid ${gridAttrs.join(" ")}>
  <row name="result" id="${rowId}">
${cells}
  </row>
</grid>`;
}

/**
 * Collect all attributes from a FetchXML query into column configs
 * Traverses root entity and all link-entities recursively
 */
export function collectAttributesFromFetchXml(
	fetchQuery: FetchNode,
	attributeTypeMap?: Map<string, string>
): LayoutColumn[] {
	const columns: LayoutColumn[] = [];

	// Helper to get the column key for an attribute
	const getColumnKey = (attr: AttributeNode, linkAlias?: string): string => {
		// If attribute has an alias, use it
		if (attr.alias) return attr.alias;
		// If from link-entity with alias, prefix with alias
		if (linkAlias) return `${linkAlias}.${attr.name}`;
		// Otherwise just the attribute name
		return attr.name;
	};

	// Collect from root entity
	if (fetchQuery.entity.attributes) {
		for (const attr of fetchQuery.entity.attributes) {
			const columnKey = getColumnKey(attr);
			const attrType = attributeTypeMap?.get(attr.name);
			columns.push({
				name: columnKey,
				width: getDefaultWidthForType(attrType),
			});
		}
	}

	// Recursive function to collect from link-entities
	const collectFromLinkEntity = (link: LinkEntityNode) => {
		// Determine the alias for this link-entity
		const linkAlias = link.alias || link.name;

		// Collect attributes
		if (link.attributes) {
			for (const attr of link.attributes) {
				const columnKey = getColumnKey(attr, linkAlias);
				const attrType = attributeTypeMap?.get(`${link.name}.${attr.name}`);
				columns.push({
					name: columnKey,
					width: getDefaultWidthForType(attrType),
					linkEntityAlias: linkAlias,
				});
			}
		}

		// Recurse into nested link-entities
		if (link.links) {
			for (const nestedLink of link.links) {
				collectFromLinkEntity(nestedLink);
			}
		}
	};

	// Collect from all link-entities
	if (fetchQuery.entity.links) {
		for (const link of fetchQuery.entity.links) {
			collectFromLinkEntity(link);
		}
	}

	return columns;
}

/**
 * Generate a LayoutXmlConfig from a FetchXML query
 * Used when no layoutxml exists (custom FetchXML)
 */
export function generateLayoutFromFetchXml(
	fetchQuery: FetchNode,
	attributeTypeMap?: Map<string, string>
): LayoutXmlConfig {
	const entityName = fetchQuery.entity.name;
	const columns = collectAttributesFromFetchXml(fetchQuery, attributeTypeMap);

	return {
		gridName: "resultset",
		jumpAttribute: columns.length > 0 ? columns[0].name : undefined,
		enableSelection: true,
		showIcon: true,
		enablePreview: true,
		primaryIdAttribute: `${entityName}id`,
		columns,
	};
}

/**
 * Merge existing layout with FetchXML changes
 * - Keeps existing column widths/order for columns that still exist
 * - Adds new columns from FetchXML at the end
 * - Removes columns no longer in FetchXML
 */
export function mergeLayoutWithFetchXml(
	existingConfig: LayoutXmlConfig,
	fetchQuery: FetchNode,
	attributeTypeMap?: Map<string, string>
): LayoutXmlConfig {
	const fetchColumns = collectAttributesFromFetchXml(fetchQuery, attributeTypeMap);
	const fetchColumnNames = new Set(fetchColumns.map((c) => c.name));
	const existingColumnNames = new Set(existingConfig.columns.map((c) => c.name));

	// Keep existing columns that are still in FetchXML (preserves order and width)
	const mergedColumns: LayoutColumn[] = existingConfig.columns.filter((col) =>
		fetchColumnNames.has(col.name)
	);

	// Add new columns from FetchXML that weren't in existing layout
	for (const col of fetchColumns) {
		if (!existingColumnNames.has(col.name)) {
			mergedColumns.push(col);
		}
	}

	return {
		...existingConfig,
		columns: mergedColumns,
	};
}

/**
 * Check if a LayoutXmlConfig matches the FetchXML attributes
 * Returns true if all FetchXML attributes are represented in the layout
 */
export function isLayoutValidForFetchXml(config: LayoutXmlConfig, fetchQuery: FetchNode): boolean {
	const fetchColumns = collectAttributesFromFetchXml(fetchQuery);
	const layoutColumnNames = new Set(config.columns.map((c) => c.name));

	// Check if all FetchXML attributes are in the layout
	for (const col of fetchColumns) {
		if (!layoutColumnNames.has(col.name)) {
			return false;
		}
	}

	return true;
}

/**
 * Update a single column's width in the config
 */
export function updateColumnWidth(
	config: LayoutXmlConfig,
	columnName: string,
	newWidth: number
): LayoutXmlConfig {
	return {
		...config,
		columns: config.columns.map((col) =>
			col.name === columnName ? { ...col, width: newWidth } : col
		),
	};
}

/**
 * Reorder columns in the config
 */
export function reorderColumns(
	config: LayoutXmlConfig,
	fromIndex: number,
	toIndex: number
): LayoutXmlConfig {
	const columns = [...config.columns];
	const [removed] = columns.splice(fromIndex, 1);
	columns.splice(toIndex, 0, removed);

	return {
		...config,
		columns,
	};
}
