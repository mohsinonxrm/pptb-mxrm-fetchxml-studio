/**
 * Utilities for validating and sanitizing FetchXML alias values.
 *
 * Dataverse alias rules:
 *   - Allowed characters: [A-Za-z0-9_]
 *   - First character must be a letter or underscore: [A-Za-z_]
 */

const ALIAS_INVALID_CHARS = /[^A-Za-z0-9_]/g;
const ALIAS_REGEX = /^[A-Za-z_][A-Za-z0-9_]*$/;

/**
 * Returns true when the alias is either empty (no alias) or matches the
 * Dataverse alias character rules.
 */
export function isValidAlias(alias: string): boolean {
	if (!alias) return true;
	return ALIAS_REGEX.test(alias);
}

/**
 * Strips characters not allowed in a Dataverse alias and fixes a leading
 * digit by prepending an underscore.  Returns an empty string when the
 * input was empty or consisted entirely of invalid characters.
 */
export function sanitizeAlias(raw: string): string {
	// Remove every character that is not [A-Za-z0-9_]
	let sanitized = raw.replace(ALIAS_INVALID_CHARS, "");

	// First character must not be a digit
	if (sanitized.length > 0 && /^\d/.test(sanitized)) {
		sanitized = `_${sanitized}`;
	}

	return sanitized;
}
