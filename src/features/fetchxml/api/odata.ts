/**
 * OData query building (#81): values placed inside single-quoted OData string literals have their
 * quotes doubled, as the OData URL conventions require, so a value can't end the literal early.
 * https://docs.oasis-open.org/odata/odata/v4.01/odata-v4.01-part2-url-conventions.html#sec_URLSyntax
 *
 * Values aren't URL-encoded here: the host builds the request URL from the query.
 */
export function odataString(value: string): string {
	return value.replace(/'/g, "''");
}
