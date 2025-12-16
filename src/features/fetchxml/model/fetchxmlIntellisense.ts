/**
 * FetchXML Intellisense Provider for Monaco Editor
 * Provides autocomplete suggestions for FetchXML elements, attributes, and values
 * Based on the official Fetch.xsd schema
 */

import type * as Monaco from "monaco-editor";

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

/**
 * FetchXML element definitions with their allowed attributes and children
 * Derived from Fetch.xsd schema
 */
const FETCHXML_SCHEMA: { elements: Record<string, ElementDef> } = {
	elements: {
		fetch: {
			description: "Root element for a FetchXML query",
			attributes: {
				version: { description: "FetchXML version" },
				count: { description: "Number of records per page", type: "integer" },
				page: { description: "Page number to retrieve", type: "integer" },
				"paging-cookie": { description: "Paging cookie for subsequent pages" },
				"utc-offset": { description: "UTC offset in minutes", type: "integer" },
				aggregate: { description: "Enable aggregate queries", type: "boolean" },
				distinct: { description: "Return only distinct records", type: "boolean" },
				top: { description: "Maximum number of records to return", type: "integer" },
				mapping: { description: "Data mapping type", values: ["internal", "logical"] },
				"min-active-row-version": { description: "Minimum active row version", type: "boolean" },
				"output-format": {
					description: "Output format",
					values: ["xml-ado", "xml-auto", "xml-elements", "xml-raw", "xml-platform"],
				},
				returntotalrecordcount: { description: "Return total record count", type: "boolean" },
				"no-lock": { description: "Disable row locking", type: "boolean" },
			},
			children: ["entity", "order"],
		},
		entity: {
			description: "Primary entity for the query",
			attributes: {
				name: { description: "Logical name of the entity", required: true },
				enableprefiltering: { description: "Enable pre-filtering for reports", type: "boolean" },
				prefilterparametername: { description: "Pre-filter parameter name" },
			},
			children: ["all-attributes", "attribute", "order", "filter", "link-entity"],
		},
		"link-entity": {
			description: "Join to a related entity",
			attributes: {
				name: { description: "Logical name of the linked entity", required: true },
				from: { description: "Attribute on the linked entity to join from" },
				to: { description: "Attribute on the parent entity to join to" },
				alias: { description: "Alias for the linked entity" },
				"link-type": {
					description: "Type of join",
					values: [
						"inner",
						"outer",
						"any",
						"not any",
						"all",
						"not all",
						"exists",
						"in",
						"matchfirstrowusingcrossapply",
					],
				},
				visible: { description: "Include in results", type: "boolean" },
				intersect: { description: "Used for N:N relationships", type: "boolean" },
				enableprefiltering: { description: "Enable pre-filtering for reports", type: "boolean" },
				prefilterparametername: { description: "Pre-filter parameter name" },
			},
			children: ["all-attributes", "attribute", "order", "filter", "link-entity"],
		},
		"all-attributes": {
			description: "Select all attributes from the entity",
			attributes: {},
			children: [],
		},
		attribute: {
			description: "Select a specific attribute",
			attributes: {
				name: { description: "Logical name of the attribute", required: true },
				alias: { description: "Alias for the attribute in results" },
				aggregate: {
					description: "Aggregate function",
					values: ["count", "countcolumn", "sum", "avg", "min", "max"],
				},
				groupby: { description: "Group by this attribute", type: "boolean" },
				dategrouping: {
					description: "Date grouping",
					values: ["day", "week", "month", "quarter", "year", "fiscal-period", "fiscal-year"],
				},
				usertimezone: { description: "Use user timezone", type: "boolean" },
				distinct: { description: "Distinct values only", type: "boolean" },
			},
			children: [],
		},
		order: {
			description: "Sort order specification",
			attributes: {
				attribute: { description: "Attribute to sort by" },
				alias: { description: "Alias of attribute to sort by (for aggregates)" },
				descending: { description: "Sort in descending order", type: "boolean" },
			},
			children: [],
		},
		filter: {
			description: "Filter criteria (WHERE clause)",
			attributes: {
				type: { description: "Logical operator", values: ["and", "or"], default: "and" },
				isquickfindfields: { description: "Quick find filter", type: "boolean" },
			},
			children: ["condition", "filter", "link-entity"],
		},
		condition: {
			description: "Filter condition",
			attributes: {
				attribute: { description: "Attribute to filter on", required: true },
				operator: {
					description: "Comparison operator",
					required: true,
					values: [
						// Comparison
						"eq",
						"ne",
						"neq",
						"gt",
						"ge",
						"lt",
						"le",
						// String
						"like",
						"not-like",
						"begins-with",
						"not-begin-with",
						"ends-with",
						"not-end-with",
						// Collection
						"in",
						"not-in",
						"between",
						"not-between",
						"contain-values",
						"not-contain-values",
						// Null
						"null",
						"not-null",
						// Date relative
						"yesterday",
						"today",
						"tomorrow",
						"last-seven-days",
						"next-seven-days",
						"last-week",
						"this-week",
						"next-week",
						"last-month",
						"this-month",
						"next-month",
						"last-year",
						"this-year",
						"next-year",
						"on",
						"on-or-before",
						"on-or-after",
						// Date X
						"last-x-hours",
						"next-x-hours",
						"last-x-days",
						"next-x-days",
						"last-x-weeks",
						"next-x-weeks",
						"last-x-months",
						"next-x-months",
						"last-x-years",
						"next-x-years",
						"olderthan-x-minutes",
						"olderthan-x-hours",
						"olderthan-x-days",
						"olderthan-x-weeks",
						"olderthan-x-months",
						"olderthan-x-years",
						// Fiscal
						"this-fiscal-year",
						"this-fiscal-period",
						"next-fiscal-year",
						"next-fiscal-period",
						"last-fiscal-year",
						"last-fiscal-period",
						"last-x-fiscal-years",
						"last-x-fiscal-periods",
						"next-x-fiscal-years",
						"next-x-fiscal-periods",
						"in-fiscal-year",
						"in-fiscal-period",
						"in-fiscal-period-and-year",
						"in-or-before-fiscal-period-and-year",
						"in-or-after-fiscal-period-and-year",
						// User/Business
						"eq-userid",
						"ne-userid",
						"eq-userteams",
						"eq-useroruserteams",
						"eq-useroruserhierarchy",
						"eq-useroruserhierarchyandteams",
						"eq-businessid",
						"ne-businessid",
						"eq-userlanguage",
						// Hierarchy
						"under",
						"eq-or-under",
						"not-under",
						"above",
						"eq-or-above",
					],
				},
				value: { description: "Value to compare against" },
				valueof: { description: "Compare to another attribute value" },
				entityname: { description: "Entity alias for link-entity conditions" },
				aggregate: {
					description: "Aggregate function for HAVING clause",
					values: ["count", "countcolumn", "sum", "avg", "min", "max"],
				},
				alias: { description: "Alias for the condition" },
				uiname: { description: "UI display name" },
				uitype: { description: "UI type" },
				uihidden: { description: "Hide in UI", values: ["0", "1"] },
			},
			children: ["value"],
		},
		value: {
			description: "Value element for multi-value operators (in, not-in, etc.)",
			attributes: {
				uiname: { description: "UI display name" },
				uitype: { description: "UI type" },
			},
			children: [],
		},
	},
} as const;

type ElementName = keyof typeof FETCHXML_SCHEMA.elements;

/**
 * Get the current element context from cursor position
 */
function getElementContext(
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
	const attrValueMatch = textBeforeCursor.match(/(\w+)\s*=\s*["']([^"']*)$/);
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
