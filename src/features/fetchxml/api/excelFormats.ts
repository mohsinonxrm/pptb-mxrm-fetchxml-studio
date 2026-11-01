/**
 * Number formats for the local Excel export (#184)
 *
 * Money keeps the record's own currency: its symbol and decimals come from Dataverse's formatted
 * value ("€1.234,50", "¥1,235", "1 234,50 kr"), not a fixed "$". Excel then shows the number with
 * its own separators.
 */

/** "#,##0" or "#,##0.00…" */
export function groupedNumberFormat(decimals: number): string {
	return decimals > 0 ? `#,##0.${"0".repeat(decimals)}` : "#,##0";
}

/** Decimal and Double: the column's precision ("0" for none, never "0.") */
export function decimalNumberFormat(precision: number | undefined): string {
	const decimals = precision ?? 2;
	return decimals > 0 ? `0.${"0".repeat(decimals)}` : "0";
}

/**
 * How many decimals a formatted amount shows. The last "." or "," is the decimal separator unless
 * it groups thousands: three digits after it, no other separator before it, and a whole number.
 */
export function decimalsShown(formatted: string, raw: number): number | undefined {
	const m = /([.,])(\d+)\D*$/.exec(formatted);
	if (!m) return /\d/.test(formatted) ? 0 : undefined;
	const [, separator, digits] = m;
	if (digits.length !== 3) return digits.length;
	const before = formatted.slice(0, m.index);
	const otherSeparator = separator === "." ? "," : ".";
	const groupsThousands = !before.includes(otherSeparator) && Number.isInteger(raw);
	return groupsThousands ? 0 : 3;
}

/**
 * The money format for one cell: the formatted value's currency symbol, before or after the
 * number as Dataverse put it, with its decimals. Plain "#,##0.00" when there's no formatted value.
 */
export function moneyNumberFormat(formatted: unknown, raw: number, precision?: number): string {
	const text = typeof formatted === "string" ? formatted.trim() : "";
	const decimals = precision ?? (text ? decimalsShown(text, raw) : undefined) ?? 2;
	const number = groupedNumberFormat(decimals);
	if (!text) return number;
	// The symbol and any space next to it; signs and brackets of negative amounts aren't part of it
	const unsigned = text.replace(/^[-−(\s]+|[)\s]+$/g, "");
	const prefix = /^[^\d]+/.exec(unsigned)?.[0] ?? "";
	const suffix = /[^\d]+$/.exec(unsigned)?.[0] ?? "";
	const literal = (s: string) => `"${s.replace(/"/g, "")}"`;
	if (prefix.trim()) return `${literal(prefix)}${number}`;
	if (suffix.trim()) return `${number}${literal(suffix)}`;
	return number;
}
