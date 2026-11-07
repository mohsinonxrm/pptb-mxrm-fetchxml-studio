/**
 * Wildcard helper for string conditions (#98): builds a like / not-like pattern from a match
 * choice, escapes wildcard characters for a literal match, and reads a stored value back.
 * https://learn.microsoft.com/power-apps/developer/data-platform/wildcard-characters?tabs=fetchxml
 *
 * - % (any characters), _ (one character) and [ ] / [^ ] (a set) are wildcards in like, not-like,
 *   begins-with, not-begin-with, ends-with and not-end-with values, and in the Web API's contains,
 *   startswith and endswith.
 * - "To use a wildcard character as a literal character, enclose the wildcard character in
 *   brackets": % → [%], _ → [_], [ → [[]. Verified live (2026-10-17): like "%[%]%", "%[_]%",
 *   "%[[]%" and begins-with "[%]" run and match only the literal character.
 * - begins-with / ends-with add their own wildcard, so their value is the text alone.
 */

/** Operators whose value is a pattern the user writes in full */
export const LIKE_OPERATORS: ReadonlySet<string> = new Set(["like", "not-like"]);
/** Operators that add the wildcard themselves (begins-with adds a trailing %, ends-with a leading %) */
export const AFFIX_OPERATORS: ReadonlySet<string> = new Set(["begins-with", "not-begin-with", "ends-with", "not-end-with"]);

/** How a like / not-like value matches: %text%, text%, %text, or the pattern as written */
export type LikeMatch = "contains" | "begins" | "ends" | "custom";

export interface WildcardValue {
	/** like / not-like only; "custom" for the other operators */
	match: LikeMatch;
	/** The text to find (the whole pattern for "custom") */
	text: string;
	/** % _ [ in the text are matched as characters, not wildcards */
	literal: boolean;
}

const WILDCARD = /[%_[]/;
const ESCAPED = /\[[%_[]\]/g;

/** Escapes the wildcard characters so they match literally: [ → [[], % → [%], _ → [_]. */
export function escapeWildcards(text: string): string {
	return text.replace(/[%_[]/g, (c) => `[${c}]`);
}

/** Reverses escapeWildcards: [[] → [, [%] → %, [_] → _. */
export function unescapeWildcards(pattern: string): string {
	return pattern.replace(ESCAPED, (s) => s[1]);
}

/** Whether the text has a wildcard character (%, _ or [) */
export const hasWildcard = (text: string): boolean => WILDCARD.test(text);

/** Whether a pattern has a wildcard that isn't one of the bracket escapes */
const hasActiveWildcard = (pattern: string): boolean => WILDCARD.test(pattern.replace(ESCAPED, ""));

/** The value to store for the operator. */
export function buildWildcardValue(operator: string, { match, text, literal }: WildcardValue): string {
	if (LIKE_OPERATORS.has(operator) && match === "custom") return text;
	const t = literal ? escapeWildcards(text) : text;
	if (!LIKE_OPERATORS.has(operator)) return t;
	if (match === "contains") return `%${t}%`;
	if (match === "begins") return `${t}%`;
	return `%${t}`;
}

/**
 * Reads a stored value back into the helper's choices. For like / not-like, %text% opens as
 * Contains, text% as Begins with and %text as Ends with when the text has no other wildcard;
 * anything else opens as Custom pattern. Bracket escapes in the text turn Match literally on.
 */
export function parseWildcardValue(operator: string, value: unknown): WildcardValue {
	const raw = typeof value === "string" ? value : value === undefined || value === null ? "" : String(value);
	const plain = (text: string) => !hasActiveWildcard(text);
	const read = (text: string, match: LikeMatch): WildcardValue => {
		const literal = text !== unescapeWildcards(text);
		return { match, text: literal ? unescapeWildcards(text) : text, literal };
	};
	if (!LIKE_OPERATORS.has(operator)) {
		return plain(raw) ? read(raw, "custom") : { match: "custom", text: raw, literal: false };
	}
	// A new like condition starts as Contains
	if (raw === "") return { match: "contains", text: "", literal: false };
	const leading = raw.startsWith("%");
	// A trailing % that isn't the end of a [%] escape
	const trailing = raw.endsWith("%") && !raw.endsWith("[%]");
	if (leading && trailing && raw.length > 2 && plain(raw.slice(1, -1))) return read(raw.slice(1, -1), "contains");
	if (trailing && !leading && raw.length > 1 && plain(raw.slice(0, -1))) return read(raw.slice(0, -1), "begins");
	if (leading && !trailing && raw.length > 1 && plain(raw.slice(1))) return read(raw.slice(1), "ends");
	return { match: "custom", text: raw, literal: false };
}

/** Whether the helper applies to the operator */
export const isWildcardOperator = (operator: string): boolean => LIKE_OPERATORS.has(operator) || AFFIX_OPERATORS.has(operator);

/**
 * The value after an operator change between wildcard operators: the text and Match literally are
 * kept, and the wildcards follow the operator (like "%abc%" → begins-with "abc", begins-with "abc"
 * → like "abc%"). A custom pattern is kept as written. Other operator changes keep the value.
 */
export function convertWildcardValue(fromOperator: string, toOperator: string, value: unknown): unknown {
	if (!isWildcardOperator(fromOperator) || !isWildcardOperator(toOperator) || typeof value !== "string" || value === "") return value;
	const parsed = parseWildcardValue(fromOperator, value);
	if (LIKE_OPERATORS.has(fromOperator) && parsed.match === "custom") return value;
	// begins-with / ends-with → like: the wildcard the operator added becomes part of the pattern
	const fromEnds = fromOperator === "ends-with" || fromOperator === "not-end-with";
	const match: LikeMatch = LIKE_OPERATORS.has(fromOperator) ? parsed.match : fromEnds ? "ends" : "begins";
	return buildWildcardValue(toOperator, { ...parsed, match });
}
