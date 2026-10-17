/**
 * Utilities for validating and sanitizing FetchXML alias values.
 *
 * Dataverse alias rules:
 *   - Allowed characters: [A-Za-z0-9_]
 *   - First character must be a letter or underscore: [A-Za-z_]
 */

import type { FetchNode } from "./nodes";
import { collectAllLinkEntities } from "./treeUtils";

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

/**
 * Every link-entity alias in the query, including link-entities inside filters (#94: those were
 * skipped, so a generated alias could collide with one).
 */
export function collectLinkEntityAliases(fetch: FetchNode): Set<string> {
	return new Set(collectAllLinkEntities(fetch).flatMap((link) => (link.alias ? [link.alias] : [])));
}

/**
 * Generate a unique alias for a link-entity
 * Appends a short alphanumeric suffix if the base alias already exists
 */
export function generateUniqueAlias(baseName: string, existingAliases: Set<string>): string {
	// If the base name isn't taken, use it
	if (!existingAliases.has(baseName)) {
		return baseName;
	}

	// Generate a unique suffix (e.g., contact_a1, contact_b2, etc.)
	const chars = "abcdefghijklmnopqrstuvwxyz";
	let suffix = 1;

	while (true) {
		// Generate suffixes like: _a1, _b1, ..., _z1, _a2, _b2, etc.
		const charIndex = (suffix - 1) % 26;
		const numPart = Math.floor((suffix - 1) / 26) + 1;
		const candidate = `${baseName}_${chars[charIndex]}${numPart}`;

		if (!existingAliases.has(candidate)) {
			return candidate;
		}

		suffix++;

		// Safety limit to prevent infinite loops
		if (suffix > 1000) {
			return `${baseName}_${Date.now()}`;
		}
	}
}
