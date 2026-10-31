/**
 * OData query building (#81): values placed inside single-quoted OData string literals have their
 * quotes doubled, as the OData URL conventions require, so a value can't end the literal early.
 * https://docs.oasis-open.org/odata/odata/v4.01/odata-v4.01-part2-url-conventions.html#sec_URLSyntax
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/query/filter-rows#filter-by-using-string-values
 *
 * odataString doesn't URL-encode: the host builds the request URL from the query as given, so a
 * caller that puts free text in a URL encodes it too (functionWithString does).
 */
export function odataString(value: string): string {
	return value.replace(/'/g, "''");
}

/**
 * A Web API function call with one string parameter, passed as a parameter alias (#176):
 * `Name(Param=@p1)?@p1='…'`, with the value's quotes doubled and the value URL-encoded. In the query
 * string, `& # % + / ?` can't end or change the value, and an apostrophe can't end the literal.
 * Built here rather than by the host's execute(), because not every host escapes the value (the
 * PPTB VS Code extension puts it in the URL raw, #173).
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/use-web-api-functions#passing-parameters-to-a-function
 */
export function functionWithString(functionName: string, parameterName: string, value: string): string {
	return `${functionName}(${parameterName}=@p1)?@p1='${encodeURIComponent(odataString(value))}'`;
}
