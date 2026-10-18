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

import type { FetchNode } from "./nodes";
import { collectQueryColumns } from "./queryColumns";

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
	/** JavaScript function of the image provider web resource (imageproviderfunctionname) */
	imageProviderFunctionName?: string;
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
		const imageProviderFunctionName = cell.getAttribute("imageproviderfunctionname");

		if (name) {
			columns.push({
				name,
				width: width ? parseInt(width, 10) : DEFAULT_WIDTHS.default,
				disableSorting: disableSorting === "1",
				imageProviderName: imageProviderName || undefined,
				imageProviderFunctionName: imageProviderFunctionName || undefined,
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

function escapeXmlAttribute(value: string): string {
	return value
		.replace(/&/g, "&amp;")
		.replace(/"/g, "&quot;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;");
}

/**
 * Generate LayoutXML string from a LayoutXmlConfig
 */
export function generateLayoutXml(config: LayoutXmlConfig): string {
	const attr = (name: string, value: string | number) => `${name}="${escapeXmlAttribute(String(value))}"`;
	const cells = config.columns
		.map((col) => {
			const attrs: string[] = [attr("name", col.name), attr("width", col.width)];
			if (col.disableSorting) {
				attrs.push('disableSorting="1"');
			}
			if (col.imageProviderName) {
				attrs.push(attr("imageproviderwebresource", col.imageProviderName));
			}
			if (col.imageProviderFunctionName) {
				attrs.push(attr("imageproviderfunctionname", col.imageProviderFunctionName));
			}
			return `    <cell ${attrs.join(" ")} />`;
		})
		.join("\n");

	const gridAttrs: string[] = [attr("name", config.gridName)];
	if (config.objectTypeCode !== undefined) {
		gridAttrs.push(attr("object", config.objectTypeCode));
	}
	if (config.jumpAttribute) {
		gridAttrs.push(attr("jump", config.jumpAttribute));
	}
	gridAttrs.push(`select="${config.enableSelection ? "1" : "0"}"`);
	gridAttrs.push(`icon="${config.showIcon ? "1" : "0"}"`);
	gridAttrs.push(`preview="${config.enablePreview ? "1" : "0"}"`);

	// The row id is the Entity's primary key attribute; an unknown key is left out rather than
	// written as a made-up "id" attribute (#93). buildViewLayoutXml always supplies it.
	const rowAttrs = [`name="result"`];
	if (config.primaryIdAttribute) rowAttrs.push(attr("id", config.primaryIdAttribute));

	return `<grid ${gridAttrs.join(" ")}>
  <row ${rowAttrs.join(" ")}>
${cells}
  </row>
</grid>`;
}

/** The root Entity's metadata a saved view's layoutxml needs */
export interface ViewLayoutEntity {
	objectTypeCode: number;
	primaryIdAttribute: string;
	primaryNameAttribute?: string;
}

/**
 * LayoutXML to save with a view (#93). Follows the documented layout:
 * <grid name="resultset" object="{ObjectTypeCode}" jump="{PrimaryNameAttribute}" select="1"
 * icon="1" preview="1"><row name="result" id="{PrimaryIdAttribute}"><cell .../></row></grid>
 * https://learn.microsoft.com/power-apps/developer/model-driven-apps/customize-entity-views#create-views
 *  - object and the row id always come from the root Entity's metadata.
 *  - jump keeps a loaded view's own value; otherwise it is the primary name attribute (it was the
 *    first column, which can be a lookup or a link-entity column).
 *  - Without a column layout yet (e.g. only link-entity attributes), one is generated from the
 *    query instead of saving an empty layoutxml.
 */
export function buildViewLayoutXml(
	config: LayoutXmlConfig | null,
	fetchQuery: FetchNode,
	entity: ViewLayoutEntity,
): string {
	const layout = config ?? generateLayoutFromFetchXml(fetchQuery, entity.primaryIdAttribute);
	return generateLayoutXml({
		...layout,
		objectTypeCode: entity.objectTypeCode,
		primaryIdAttribute: entity.primaryIdAttribute,
		jumpAttribute: layout.jumpAttribute || entity.primaryNameAttribute || undefined,
	});
}

/**
 * Collect all attributes from a FetchXML query into column configs, keyed as Dataverse names the
 * result columns (queryColumns.ts: alias, name, "linkalias.name" or "{name}{N}.name", #114).
 */
export function collectAttributesFromFetchXml(
	fetchQuery: FetchNode,
	attributeTypeMap?: Map<string, string>,
): LayoutColumn[] {
	return collectQueryColumns(fetchQuery).map((column) => {
		const attrType = column.isRoot
			? attributeTypeMap?.get(column.attribute.name)
			: attributeTypeMap?.get(`${column.container.name}.${column.attribute.name}`);
		return {
			name: column.key,
			width: getDefaultWidthForType(attrType),
			...(column.linkAlias ? { linkEntityAlias: column.linkAlias } : {}),
		};
	});
}

/**
 * Generate a LayoutXmlConfig from a FetchXML query
 * Used when no layoutxml exists (custom FetchXML)
 * @param primaryIdAttribute - The entity's real primary key attribute from Dataverse metadata.
 *   Must be supplied by the caller from entity metadata — no guessing is performed.
 */
export function generateLayoutFromFetchXml(
	fetchQuery: FetchNode,
	primaryIdAttribute?: string,
	attributeTypeMap?: Map<string, string>,
): LayoutXmlConfig {
	const columns = collectAttributesFromFetchXml(fetchQuery, attributeTypeMap);

	return {
		gridName: "resultset",
		// jump is the primary name attribute, from metadata when the view is saved (buildViewLayoutXml)
		jumpAttribute: undefined,
		enableSelection: true,
		showIcon: true,
		enablePreview: true,
		primaryIdAttribute,
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
	attributeTypeMap?: Map<string, string>,
): LayoutXmlConfig {
	const fetchColumns = collectAttributesFromFetchXml(fetchQuery, attributeTypeMap);
	const fetchColumnNames = new Set(fetchColumns.map((c) => c.name));
	const existingColumnNames = new Set(existingConfig.columns.map((c) => c.name));

	// Keep existing columns that are still in FetchXML (preserves order and width)
	const mergedColumns: LayoutColumn[] = existingConfig.columns.filter((col) =>
		fetchColumnNames.has(col.name),
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
	newWidth: number,
): LayoutXmlConfig {
	return {
		...config,
		columns: config.columns.map((col) =>
			col.name === columnName ? { ...col, width: newWidth } : col,
		),
	};
}

/**
 * Reorder columns in the config
 */
export function reorderColumns(
	config: LayoutXmlConfig,
	fromIndex: number,
	toIndex: number,
): LayoutXmlConfig {
	const columns = [...config.columns];
	const [removed] = columns.splice(fromIndex, 1);
	columns.splice(toIndex, 0, removed);

	return {
		...config,
		columns,
	};
}
