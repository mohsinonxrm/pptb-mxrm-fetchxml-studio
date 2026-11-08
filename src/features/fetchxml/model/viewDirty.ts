/**
 * Whether the query has unsaved changes against the loaded view (#278, owner decision Q14), for the
 * app bar's "Unsaved changes" badge and New's confirmation. Pure, so it can be unit tested.
 *
 * Model-driven apps mark a view as edited when its query or its columns change, column width
 * included: "When you change the width of a column, an asterisk appears beside the view name", and
 * Save view writes both to the view (FetchXML and LayoutXML).
 * https://learn.microsoft.com/power-apps/user/grid-filters-advanced#save-your-view
 */

import { generateFetchXml } from "./fetchxml";
import { parseFetchXml } from "./fetchxmlParser";
import { generateLayoutFromFetchXml, generateLayoutXml, getDefaultWidthForType } from "./layoutxml";

/** What changed: the FetchXML, the columns (added, removed or reordered), or a column's width */
export type ViewDirtyReason = "query" | "columns" | "widths";

export interface ViewDirtyInput {
	/** The loaded view's FetchXML (stored or generated); null when no view is loaded */
	baselineFetchXml: string | null;
	/** The query now: generated from the tree, or the editor's text */
	currentFetchXml: string;
	/** The loaded view's LayoutXML; empty, null or undefined when the view has none (one is generated from baselineFetchXml) */
	baselineLayoutXml?: string | null;
	/** The column layout now (generateLayoutXml of the column config); columns aren't compared without it */
	currentLayoutXml?: string | null;
	/** Attribute types by logical name, for the default widths of a generated baseline layout; pass the map the store generated the current layout with (none today) */
	attributeTypeMap?: Map<string, string>;
}

export interface ViewDirtyResult {
	dirty: boolean;
	/** In the order query, columns, widths */
	reasons: ViewDirtyReason[];
}

/** XML with the whitespace between tags removed and other runs collapsed, when it can't be parsed */
const collapseWhitespace = (xml: string): string =>
	xml.trim().replace(/>\s+</g, "><").replace(/\s+/g, " ");

const escapeXml = (text: string): string =>
	text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * <fetch> attributes that say which page to read, not what the query is: a view saves without them,
 * and running or paging the query mustn't mark it changed. count and top stay: they're part of the
 * query. https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/page-results
 */
const PAGING_ATTRIBUTES = new Set(["page", "paging-cookie"]);

function canonicalElement(element: Element, isRoot: boolean): string {
	const attributes = Array.from(element.attributes)
		.filter(
			(attribute) =>
				!(isRoot && element.tagName === "fetch" && PAGING_ATTRIBUTES.has(attribute.name)),
		)
		.map((attribute) => `${attribute.name}="${escapeXml(attribute.value)}"`)
		.sort()
		.join(" ");
	const nodes = Array.from(element.childNodes);
	// Whitespace between child elements is formatting; text in an element without child elements
	// (a <value>) is data and kept exactly, spaces included
	const hasElementChildren = nodes.some((node) => node.nodeType === node.ELEMENT_NODE);
	const children = nodes
		.map((node) => {
			if (node.nodeType === node.ELEMENT_NODE)
				return canonicalElement(node as Element, false);
			if (node.nodeType !== node.TEXT_NODE && node.nodeType !== node.CDATA_SECTION_NODE)
				return "";
			const text = node.textContent ?? "";
			return hasElementChildren && !text.trim() ? "" : escapeXml(text);
		})
		.join("");
	return `<${element.tagName}${attributes ? ` ${attributes}` : ""}>${children}</${element.tagName}>`;
}

/**
 * One line of XML with every element's attributes sorted by name, double quotes, whitespace between
 * elements and comments dropped, and the root <fetch>'s page and paging-cookie left out. Text inside
 * an element (a condition <value>) is kept as is. XML that isn't well-formed is compared with its
 * whitespace collapsed.
 */
function canonicalXml(xml: string): string {
	const doc = new DOMParser().parseFromString(xml, "text/xml");
	if (doc.getElementsByTagName("parsererror").length > 0 || !doc.documentElement)
		return collapseWhitespace(xml);
	return canonicalElement(doc.documentElement, true);
}

/**
 * FetchXML in the form the tree generates (parseFetchXml, then generateFetchXml), then canonical
 * (attributes sorted, paging attributes dropped), so attribute order, quotes, whitespace between
 * elements and defaults the generator leaves out (distinct="false", descending="false") aren't
 * changes. The view's stored FetchXML and the generated query only compare equal through the
 * generator (savedViews.ts isUnmodifiedView, #93); the generator keeps the <fetch> attributes it
 * passes through (version, output-format, mapping) in their original order, hence the sort.
 * Removing or adding one of those still counts as a change: they're in the text Save writes.
 * Whitespace inside a condition value counts too: it's part of the value. FetchXML the parser
 * can't build a tree from is compared canonical.
 */
export function normalizeFetchXml(fetchXml: string): string {
	if (!fetchXml.trim()) return "";
	const result = parseFetchXml(fetchXml);
	return canonicalXml(
		result.success && result.fetchNode ? generateFetchXml(result.fetchNode) : fetchXml,
	);
}

interface LayoutCell {
	name: string;
	/** The width attribute as written; the default width when it's missing, as parseLayoutXml reads it */
	width: string;
}

/**
 * The cells of a LayoutXML in order, read as parseLayoutXml reads them (a <grid> is required, cells
 * without a name are skipped, a missing width is the default) but keeping the width text, so a
 * width that isn't a number can still be compared. Null when it can't be read.
 */
function readCells(layoutXml: string): LayoutCell[] | null {
	const doc = new DOMParser().parseFromString(layoutXml, "text/xml");
	if (
		doc.getElementsByTagName("parsererror").length > 0 ||
		doc.getElementsByTagName("grid").length === 0
	)
		return null;
	return Array.from(doc.getElementsByTagName("cell")).flatMap((cell) => {
		const name = cell.getAttribute("name");
		return name
			? [{ name, width: cell.getAttribute("width") ?? String(getDefaultWidthForType()) }]
			: [];
	});
}

const NUMERIC = /^\s*\d+(\.\d+)?\s*$/;

/** Two widths are the same: equal numbers when both are numeric, else the same text */
const sameWidth = (a: string, b: string): boolean =>
	NUMERIC.test(a) && NUMERIC.test(b) ? Number(a) === Number(b) : a.trim() === b.trim();

/**
 * Column changes between two LayoutXMLs: "columns" when the cell names differ in set or order,
 * "widths" when a cell in both has another width. Only the cells' name, order and width count:
 * Model-driven apps' column edits change these. A LayoutXML that can't be read is compared as text.
 */
function layoutReasons(baseline: string, current: string): ViewDirtyReason[] {
	const before = readCells(baseline);
	const after = readCells(current);
	if (!before || !after)
		return collapseWhitespace(baseline) === collapseWhitespace(current) ? [] : ["columns"];

	const reasons: ViewDirtyReason[] = [];
	const names = (cells: LayoutCell[]) => cells.map((cell) => cell.name).join("\n");
	if (names(before) !== names(after)) reasons.push("columns");

	const beforeWidths = new Map<string, string>();
	for (const cell of before)
		if (!beforeWidths.has(cell.name)) beforeWidths.set(cell.name, cell.width);
	const widthChanged = after.some((cell) => {
		const width = beforeWidths.get(cell.name);
		return width !== undefined && !sameWidth(width, cell.width);
	});
	if (widthChanged) reasons.push("widths");
	return reasons;
}

/**
 * The layout FetchXML Studio gives a query without a stored LayoutXML (builderStore uses
 * generateLayoutFromFetchXml for it), so a view saved without one still has a baseline. Null when
 * the FetchXML can't be parsed into a tree.
 */
function generatedLayoutXml(
	fetchXml: string,
	attributeTypeMap: Map<string, string> | undefined,
): string | null {
	const result = parseFetchXml(fetchXml);
	if (!result.success || !result.fetchNode) return null;
	return generateLayoutXml(
		generateLayoutFromFetchXml(result.fetchNode, undefined, attributeTypeMap),
	);
}

/**
 * Whether the view has unsaved changes (#278, Q14):
 *  - "query": the normalised FetchXML differs, or there's a query and no view is loaded;
 *  - "columns": a column was added, removed or moved;
 *  - "widths": a column's width changed.
 * Columns are compared when there's a current LayoutXML and a view is loaded. A view without a
 * stored LayoutXML is compared with the layout generated from its FetchXML.
 */
export function isViewDirty({
	baselineFetchXml,
	currentFetchXml,
	baselineLayoutXml,
	currentLayoutXml,
	attributeTypeMap,
}: ViewDirtyInput): ViewDirtyResult {
	const reasons: ViewDirtyReason[] = [];
	const current = normalizeFetchXml(currentFetchXml);
	if (
		baselineFetchXml === null ? current !== "" : normalizeFetchXml(baselineFetchXml) !== current
	)
		reasons.push("query");
	if (currentLayoutXml?.trim() && baselineFetchXml !== null) {
		const baselineLayout = baselineLayoutXml?.trim()
			? baselineLayoutXml
			: generatedLayoutXml(baselineFetchXml, attributeTypeMap);
		if (baselineLayout) reasons.push(...layoutReasons(baselineLayout, currentLayoutXml));
	}
	return { dirty: reasons.length > 0, reasons };
}
