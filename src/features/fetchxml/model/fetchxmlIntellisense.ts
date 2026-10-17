/**
 * FetchXML Intellisense Provider for Monaco Editor
 * Provides autocomplete suggestions for FetchXML elements, attributes, and values
 * Based on the official Fetch.xsd schema
 */

import type * as Monaco from "monaco-editor";
import { KNOWN_OPERATORS } from "./operators";

/**
 * Attribute definition types
 */
interface AttributeDef {
	description: string;
	type?: "boolean" | "integer" | string;
	values?: string[];
	required?: boolean;
	default?: string;
}

interface ElementDef {
	description: string;
	attributes: Record<string, AttributeDef>;
	children: string[];
}

const LEGACY = "Not in the FetchXML reference; accepted by Dataverse (verified).";
const SAVED_VIEW = "Written by Advanced Find / saved views; Dataverse ignores it.";
const QUERY_HINTS = [
	"ForceOrder",
	"DisableRowGoal",
	"EnableOptimizerHotfixes",
	"LoopJoin",
	"MergeJoin",
	"HashJoin",
	"NO_PERFORMANCE_SPOOL",
	"ENABLE_HIST_AMENDMENT_FOR_ASC_KEYS",
];

/**
 * FetchXML element definitions with their allowed attributes and children, synced with the FetchXML
 * reference (https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/).
 * Undocumented attributes are listed only when verified against Dataverse, and say so (#94).
 */
export const FETCHXML_SCHEMA: { elements: Record<string, ElementDef> } = {
	elements: {
		fetch: {
			description: "Root element for a FetchXML query",
			attributes: {
				aggregate: { description: "Return aggregate values", type: "boolean" },
				aggregatelimit: { description: "Aggregate record limit below the standard 50,000", type: "integer" },
				count: { description: "Number of records per page", type: "integer" },
				datasource: { description: "'retained' queries long-term retained rows only", values: ["retained"] },
				distinct: { description: "Leave out duplicate rows", type: "boolean" },
				latematerialize: { description: "Break the query into smaller parts (performance)", type: "boolean" },
				"no-lock": { description: "Legacy setting to prevent shared locks; no longer necessary", type: "boolean" },
				options: {
					description: "SQL Server query hints, comma-separated (only when Microsoft support recommends them)",
					values: QUERY_HINTS,
				},
				page: { description: "Page number to return", type: "integer" },
				"paging-cookie": { description: "Paging cookie from the previous page" },
				returntotalrecordcount: { description: "Return the total number of matching records", type: "boolean" },
				top: {
					description: "Number of records to return (at most 5,000; not with page, count or returntotalrecordcount)",
					type: "integer",
				},
				useraworderby: { description: "Sort choice columns by their integer value instead of the label", type: "boolean" },
				"utc-offset": { description: `UTC offset in minutes. ${LEGACY}`, type: "integer" },
				version: { description: LEGACY },
				mapping: { description: LEGACY, values: ["logical", "internal"] },
				"output-format": { description: LEGACY, values: ["xml-platform"] },
				"min-active-row-version": { description: LEGACY, type: "boolean" },
			},
			children: ["entity"],
		},
		entity: {
			description: "The entity the query is based on (one per query)",
			attributes: {
				name: { description: "Logical name of the entity", required: true },
				enableprefiltering: { description: "Report prefiltering (SQL Server Reporting Services reports)", type: "boolean" },
				prefilterparametername: { description: "Report prefiltering parameter name (start it with CRM_)" },
			},
			children: ["all-attributes", "attribute", "order", "filter", "link-entity"],
		},
		"link-entity": {
			description: "Join a related entity, or filter on related records inside a filter",
			attributes: {
				name: { description: "Logical name of the related entity", required: true },
				from: { description: "Attribute of the related entity that matches 'to' (usually set)" },
				to: { description: "Attribute of the parent entity that matches 'from' (usually set)" },
				alias: { description: "Name to reference this link-entity in conditions, valueof and orders" },
				"link-type": {
					description:
						"inner (default), outer, exists, in, matchfirstrowusingcrossapply; inside a filter: any, not any, all, not all. natural: " +
						LEGACY,
					values: ["inner", "outer", "exists", "in", "matchfirstrowusingcrossapply", "any", "not any", "all", "not all", "natural"],
				},
				intersect: {
					description: "Shows the link-entity only joins and returns no attributes (doesn't change how the query runs)",
					type: "boolean",
				},
				visible: { description: `${LEGACY} Dataverse ignores it.`, type: "boolean" },
				enableprefiltering: { description: `Report prefiltering. ${LEGACY}`, type: "boolean" },
				prefilterparametername: { description: `Report prefiltering parameter name. ${LEGACY}` },
			},
			children: ["all-attributes", "attribute", "order", "filter", "link-entity"],
		},
		"all-attributes": {
			description: "Return every attribute with a value (not recommended for most queries)",
			attributes: {},
			children: [],
		},
		attribute: {
			description: "An attribute to return",
			attributes: {
				name: { description: "Logical name of the attribute", required: true },
				alias: { description: "Name of the returned column (required in aggregate queries)" },
				aggregate: {
					description: "Aggregate function",
					values: ["count", "countcolumn", "sum", "avg", "min", "max"],
				},
				groupby: { description: "Group by this attribute (aggregate queries)", type: "boolean" },
				dategrouping: {
					description: "Date part to group by",
					values: ["day", "week", "month", "quarter", "year", "fiscal-period", "fiscal-year"],
				},
				usertimezone: { description: "Group dates in the user's time zone (default true); false uses UTC", type: "boolean" },
				distinct: { description: "With countcolumn, count only unique values", type: "boolean" },
				rowaggregate: { description: "CountChildren returns the number of hierarchical child records", values: ["CountChildren"] },
			},
			children: [],
		},
		order: {
			description: "Sort order",
			attributes: {
				attribute: { description: "Attribute to sort by (non-aggregate queries)" },
				alias: { description: "Attribute alias to sort by (aggregate queries)" },
				descending: { description: "Sort in descending order", type: "boolean" },
				entityname: { description: "Alias of a link-entity, to apply its sort first (orders in the entity only)" },
			},
			children: [],
		},
		filter: {
			description: "Conditions that must be true",
			attributes: {
				type: { description: "All (and) or any (or) of the conditions", values: ["and", "or"], default: "and" },
				hint: { description: "union: can improve performance of an 'or' filter (one per query)", values: ["union"] },
				isquickfindfields: { description: "Run as a quick find query", type: "boolean" },
				overridequickfindrecordlimitenabled: { description: "Apply the quick find record limit (testing)", type: "boolean" },
				overridequickfindrecordlimitdisabled: { description: "Bypass the quick find record limit (testing)", type: "boolean" },
			},
			children: ["condition", "filter", "link-entity"],
		},
		condition: {
			description: "A condition on an attribute value",
			attributes: {
				// The reference lists attribute as optional, but Dataverse rejects a condition without it (0x8004111e)
				attribute: { description: "Attribute to filter on (Dataverse requires it)", required: true },
				operator: {
					description: "Comparison operator",
					required: true,
					values: [...KNOWN_OPERATORS],
				},
				value: { description: "Value for operators that take one value (use <value> elements for several)" },
				valueof: {
					description: "Compare to another attribute: name, or alias.name for a link-entity (eq, ne, gt, ge, lt, le only)",
				},
				entityname: { description: "Alias or name of a link-entity: apply the condition after the join" },
				uiname: { description: SAVED_VIEW },
				uitype: { description: SAVED_VIEW },
				uihidden: { description: SAVED_VIEW, values: ["0", "1"] },
			},
			children: ["value"],
		},
		value: {
			description: "A value for operators that take several values (in, between, ...)",
			attributes: {
				uiname: { description: SAVED_VIEW },
				uitype: { description: SAVED_VIEW },
			},
			children: [],
		},
	},
};

type ElementName = keyof typeof FETCHXML_SCHEMA.elements;

/**
 * Get the current element context from cursor position
 */
export function getElementContext(
	model: Monaco.editor.ITextModel,
	position: Monaco.Position
): {
	element: string | null;
	inAttribute: boolean;
	attributeName: string | null;
	inValue: boolean;
} {
	const textUntilPosition = model.getValueInRange({
		startLineNumber: 1,
		startColumn: 1,
		endLineNumber: position.lineNumber,
		endColumn: position.column,
	});

	// Find the last opened element
	const elementMatches = textUntilPosition.match(/<(\/?)([\w-]+)(?:\s|>|$)/g);
	const openElements: string[] = [];

	if (elementMatches) {
		for (const match of elementMatches) {
			const isClosing = match.startsWith("</");
			const elementName = match.replace(/<\/?/, "").replace(/[\s>].*/, "");

			if (isClosing) {
				// Pop the last matching open element
				const idx = openElements.lastIndexOf(elementName);
				if (idx !== -1) {
					openElements.splice(idx, 1);
				}
			} else if (!match.includes("/>")) {
				openElements.push(elementName);
			}
		}
	}

	const currentElement = openElements[openElements.length - 1] || null;

	// Check if we're inside an attribute value
	const lineText = model.getLineContent(position.lineNumber);
	const textBeforeCursor = lineText.substring(0, position.column - 1);

	// Check if inside attribute value (after = and inside quotes)
	// Attribute names can contain hyphens (link-type, paging-cookie, no-lock, utc-offset) (#94)
	const attrValueMatch = textBeforeCursor.match(/([\w-]+)\s*=\s*["']([^"']*)$/);
	if (attrValueMatch) {
		return {
			element: currentElement,
			inAttribute: true,
			attributeName: attrValueMatch[1],
			inValue: true,
		};
	}

	// Check if we're in an opening tag (for attribute suggestions)
	const inOpeningTag = /<[\w-]+[^>]*$/.test(textBeforeCursor);

	return {
		element: currentElement,
		inAttribute: inOpeningTag,
		attributeName: null,
		inValue: false,
	};
}

/**
 * Create Monaco completion items for FetchXML
 */
function createCompletionItem(
	label: string,
	kind: Monaco.languages.CompletionItemKind,
	detail: string,
	insertText: string,
	range: Monaco.IRange,
	documentation?: string
): Monaco.languages.CompletionItem {
	return {
		label,
		kind,
		detail,
		insertText,
		range,
		documentation,
		insertTextRules: insertText.includes("$") ? 4 : undefined, // InsertAsSnippet if contains $
	};
}

/**
 * Register FetchXML completion provider with Monaco
 */
export function registerFetchXmlIntellisense(monaco: typeof Monaco): Monaco.IDisposable {
	return monaco.languages.registerCompletionItemProvider("xml", {
		triggerCharacters: ["<", " ", "=", '"', "'"],

		provideCompletionItems(
			model: Monaco.editor.ITextModel,
			position: Monaco.Position
		): Monaco.languages.ProviderResult<Monaco.languages.CompletionList> {
			const word = model.getWordUntilPosition(position);
			const range: Monaco.IRange = {
				startLineNumber: position.lineNumber,
				startColumn: word.startColumn,
				endLineNumber: position.lineNumber,
				endColumn: word.endColumn,
			};

			const context = getElementContext(model, position);
			const lineText = model.getLineContent(position.lineNumber);
			const textBeforeCursor = lineText.substring(0, position.column - 1);

			const suggestions: Monaco.languages.CompletionItem[] = [];

			// Inside attribute value - suggest valid values
			if (context.inValue && context.attributeName && context.element) {
				const elementDef = FETCHXML_SCHEMA.elements[context.element as ElementName];
				if (elementDef) {
					const attrDef =
						elementDef.attributes[context.attributeName as keyof typeof elementDef.attributes];
					if (attrDef && "values" in attrDef && attrDef.values) {
						for (const value of attrDef.values) {
							suggestions.push(
								createCompletionItem(
									value,
									monaco.languages.CompletionItemKind.Value,
									`Value for ${context.attributeName}`,
									value,
									range
								)
							);
						}
					} else if (attrDef && "type" in attrDef && attrDef.type === "boolean") {
						suggestions.push(
							createCompletionItem(
								"true",
								monaco.languages.CompletionItemKind.Value,
								"Boolean true",
								"true",
								range
							)
						);
						suggestions.push(
							createCompletionItem(
								"false",
								monaco.languages.CompletionItemKind.Value,
								"Boolean false",
								"false",
								range
							)
						);
					}
				}
				return { suggestions };
			}

			// Inside opening tag - suggest attributes
			if (context.inAttribute && context.element && !context.inValue) {
				const elementDef = FETCHXML_SCHEMA.elements[context.element as ElementName];
				if (elementDef) {
					for (const [attrName, attrDef] of Object.entries(elementDef.attributes)) {
						// Check if attribute already exists on this line
						if (lineText.includes(`${attrName}=`)) continue;

						const required = "required" in attrDef && attrDef.required;
						const hasValues = "values" in attrDef && attrDef.values;

						suggestions.push(
							createCompletionItem(
								attrName,
								monaco.languages.CompletionItemKind.Property,
								required ? "(required)" : "(optional)",
								hasValues ? `${attrName}="$1"` : `${attrName}="$1"`,
								range,
								attrDef.description
							)
						);
					}
				}
				return { suggestions };
			}

			// After < - suggest elements
			if (textBeforeCursor.endsWith("<") || textBeforeCursor.match(/<\w*$/)) {
				// Determine valid child elements based on parent
				let validElements: string[] = [];

				if (!context.element) {
					// Root level - only fetch is valid
					validElements = ["fetch"];
				} else {
					const parentDef = FETCHXML_SCHEMA.elements[context.element as ElementName];
					if (parentDef) {
						validElements = [...parentDef.children];
					}
				}

				// Add closing tag suggestion if in an element
				if (context.element) {
					suggestions.push(
						createCompletionItem(
							`/${context.element}`,
							monaco.languages.CompletionItemKind.Keyword,
							`Close <${context.element}> tag`,
							`/${context.element}>`,
							range
						)
					);
				}

				for (const elemName of validElements) {
					const elemDef = FETCHXML_SCHEMA.elements[elemName as ElementName];
					if (elemDef) {
						// Find required attributes
						const requiredAttrs = Object.entries(elemDef.attributes)
							.filter(([, def]) => "required" in def && def.required)
							.map(([name]) => name);

						// Create snippet with required attributes
						let snippet = elemName;
						let tabStop = 1;
						for (const attr of requiredAttrs) {
							snippet += ` ${attr}="$${tabStop++}"`;
						}

						// Self-closing if no children, otherwise add closing tag
						if (elemDef.children.length === 0) {
							snippet += " />";
						} else {
							snippet += ">$0</" + elemName + ">";
						}

						suggestions.push(
							createCompletionItem(
								elemName,
								monaco.languages.CompletionItemKind.Class,
								"FetchXML element",
								snippet,
								range,
								elemDef.description
							)
						);
					}
				}

				return { suggestions };
			}

			return { suggestions };
		},
	});
}

/**
 * Get FetchXML hover information
 */
export function registerFetchXmlHoverProvider(monaco: typeof Monaco): Monaco.IDisposable {
	return monaco.languages.registerHoverProvider("xml", {
		provideHover(
			model: Monaco.editor.ITextModel,
			position: Monaco.Position
		): Monaco.languages.ProviderResult<Monaco.languages.Hover> {
			const word = model.getWordAtPosition(position);
			if (!word) return null;

			const lineText = model.getLineContent(position.lineNumber);

			// Check if hovering over an element name
			const elementMatch = lineText.match(new RegExp(`<(/?)(${word.word})(?:\\s|>|/)`));
			if (elementMatch) {
				const elemDef = FETCHXML_SCHEMA.elements[word.word as ElementName];
				if (elemDef) {
					return {
						contents: [
							{ value: `**<${word.word}>**` },
							{ value: elemDef.description },
							{
								value:
									elemDef.children.length > 0
										? `\n**Child elements:** ${elemDef.children.join(", ")}`
										: "\n*Self-closing element*",
							},
						],
					};
				}
			}

			// Check if hovering over an attribute name
			const attrMatch = lineText.match(new RegExp(`(${word.word})\\s*=`));
			if (attrMatch) {
				// Find parent element
				const context = getElementContext(model, position);
				if (context.element) {
					const elemDef = FETCHXML_SCHEMA.elements[context.element as ElementName];
					if (elemDef) {
						const attrDef = elemDef.attributes[word.word as keyof typeof elemDef.attributes];
						if (attrDef) {
							const contents: Monaco.IMarkdownString[] = [
								{ value: `**${word.word}**` },
								{ value: attrDef.description },
							];

							if ("values" in attrDef && attrDef.values) {
								contents.push({ value: `\n**Valid values:** ${attrDef.values.join(", ")}` });
							}
							if ("type" in attrDef && attrDef.type) {
								contents.push({ value: `\n**Type:** ${attrDef.type}` });
							}
							if ("required" in attrDef && attrDef.required) {
								contents.push({ value: "\n*Required attribute*" });
							}

							return { contents };
						}
					}
				}
			}

			return null;
		},
	});
}
