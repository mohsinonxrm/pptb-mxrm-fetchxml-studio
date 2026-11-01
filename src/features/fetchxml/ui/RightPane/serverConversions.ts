/**
 * When the Code tab asks Dataverse to convert the FetchXML (#199)
 *
 * The C# QueryExpression and SQL formats call FetchXmlToQueryExpression and FetchXMLToSQL. They're
 * requested only while their format is selected, once typing pauses, and once per query and
 * environment: switching formats back and forth reuses the result.
 */

/** Wait after the last change before converting, so typing a value doesn't send a request per key */
export const CONVERSION_DELAY_MS = 500;

/** The formats Dataverse converts */
export type ServerConversion = "queryexpression" | "sql";

/** The conversion a format needs, if any (the other formats are generated in the tool) */
export function serverConversionFor(format: string): ServerConversion | undefined {
	if (format === "csharp-queryexpr") return "queryexpression";
	if (format === "sql") return "sql";
	return undefined;
}

/** A conversion result is only shown for the query and environment it was made for */
export const conversionKey = (fetchXml: string, environmentUrl: string | undefined): string =>
	`${environmentUrl ?? ""}\n${fetchXml}`;

/** A finished conversion */
export interface ConversionResult {
	key: string;
	code: string;
	error: string | null;
}

/**
 * Whether to send a request for `key`: not when one was already sent for it, unless that one
 * failed (so selecting the format again retries).
 */
export function needsConversion(key: string, lastRequested: string | undefined, result: ConversionResult | undefined): boolean {
	if (lastRequested !== key) return true;
	return result?.key === key && result.error !== null;
}
