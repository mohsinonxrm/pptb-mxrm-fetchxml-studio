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
import { parseLayoutXml } from "./layoutxml";

/** What changed: the FetchXML, the columns (added, removed or reordered), or a column's width */
export type ViewDirtyReason = "query" | "columns" | "widths";

export interface ViewDirtyInput {
	/** The loaded view's FetchXML (stored or generated); null when no view is loaded */
	baselineFetchXml: string | null;
	/** The query now: generated from the tree, or the editor's text */
	currentFetchXml: string;
	/** The loaded view's LayoutXML; null or undefined when there's none to compare with */
	baselineLayoutXml?: string | null;
	/** The column layout now (generateLayoutXml of the column config) */
	currentLayoutXml?: string | null;
}

export interface ViewDirtyResult {
	dirty: boolean;
	/** In the order query, columns, widths */
	reasons: ViewDirtyReason[];
}

/** XML with the whitespace between tags removed and other runs collapsed, when it can't be parsed */
const collapseWhitespace = (xml: string): string => xml.trim().replace(/>\s+</g, "><").replace(/\s+/g, " ");

const escapeXml = (text: string): string =>
	text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function canonicalElement(element: Element): string {
	const attributes = Array.from(element.attributes)
		.map((attribute) => `${attribute.name}="${escapeXml(attribute.value)}"`)
		.sort()
		.join(" ");
	const children = Array.from(element.childNodes)
		.map((node) => {
			if (node.nodeType === node.ELEMENT_NODE) return canonicalElement(node as Element);
			if (node.nodeType === node.TEXT_NODE || node.nodeType === node.CDATA_SECTION_NODE) return escapeXml((node.textContent ?? "").trim());
			return "";
		})
		.join("");
	return `<${element.tagName}${attributes ? ` ${attributes}` : ""}>${children}</${element.tagName}>`;
}

/**
 * One line of XML with every element's attributes sorted by name, double quotes, whitespace-only
 * text and comments dropped, and text trimmed. XML that isn't well-formed is compared with its
 * whitespace collapsed.
 */
function canonicalXml(xml: string): string {
	const doc = new DOMParser().parseFromString(xml, "text/xml");
	if (doc.getElementsByTagName("parsererror").length > 0 || !doc.documentElement) return collapseWhitespace(xml);
	return canonicalElement(doc.documentElement);
}

/**
 * FetchXML in the form the tree generates (parseFetchXml, then generateFetchXml), then canonical
 * (attributes sorted), so attribute order, quotes, whitespace and defaults the generator leaves out
 * (distinct="false", descending="false") aren't changes. The view's stored FetchXML and the
 * generated query only compare equal through the generator (savedViews.ts isUnmodifiedView, #93);
 * the generator keeps the <fetch> attributes it passes through (version, mapping…) in their
 * original order, hence the sort. FetchXML the parser can't build a tree from is compared canonical.
 */
export function normalizeFetchXml(fetchXml: string): string {
	if (!fetchXml.trim()) return "";
	const result = parseFetchXml(fetchXml);
	return canonicalXml(result.success && result.fetchNode ? generateFetchXml(result.fetchNode) : fetchXml);
}

interface LayoutCell {
	name: string;
	width: number;
}

/** The cells of a LayoutXML in order (parseLayoutXml, so <cell> attribute order doesn't matter); null when it can't be read */
function readCells(layoutXml: string): LayoutCell[] | null {
	try {
		return parseLayoutXml(layoutXml).columns.map(({ name, width }) => ({ name, width }));
	} catch {
		return null;
	}
}

/**
 * Column changes between two LayoutXMLs: "columns" when the cell names differ in set or order,
 * "widths" when a cell in both has another width. Only the cells' name, order and width count:
 * Model-driven apps' column edits change these. A LayoutXML that can't be read is compared as text.
 */
function layoutReasons(baseline: string, current: string): ViewDirtyReason[] {
	const before = readCells(baseline);
	const after = readCells(current);
	if (!before || !after) return collapseWhitespace(baseline) === collapseWhitespace(current) ? [] : ["columns"];

	const reasons: ViewDirtyReason[] = [];
	const names = (cells: LayoutCell[]) => cells.map((cell) => cell.name).join("\n");
	if (names(before) !== names(after)) reasons.push("columns");

	const beforeWidths = new Map<string, number>();
	for (const cell of before) if (!beforeWidths.has(cell.name)) beforeWidths.set(cell.name, cell.width);
	const widthChanged = after.some((cell) => beforeWidths.has(cell.name) && beforeWidths.get(cell.name) !== cell.width);
	if (widthChanged) reasons.push("widths");
	return reasons;
}

/**
 * Whether the view has unsaved changes (#278, Q14):
 *  - "query": the normalised FetchXML differs, or there's a query and no view is loaded;
 *  - "columns": a column was added, removed or moved;
 *  - "widths": a column's width changed.
 * Columns are compared only when both LayoutXMLs are given.
 */
export function isViewDirty({ baselineFetchXml, currentFetchXml, baselineLayoutXml, currentLayoutXml }: ViewDirtyInput): ViewDirtyResult {
	const reasons: ViewDirtyReason[] = [];
	const current = normalizeFetchXml(currentFetchXml);
	if (baselineFetchXml === null ? current !== "" : normalizeFetchXml(baselineFetchXml) !== current) reasons.push("query");
	if (baselineLayoutXml && currentLayoutXml) reasons.push(...layoutReasons(baselineLayoutXml, currentLayoutXml));
	return { dirty: reasons.length > 0, reasons };
}
