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
 * it groups thousands. With exactly three digits after it, both readings are tried and the one
 * closer to the raw amount wins ("¥1,235" for 1235 has none, "BD 1.234" for 1.234 has three).
 */
export function decimalsShown(formatted: string, raw: number): number | undefined {
	const m = /([.,])(\d+)\D*$/.exec(formatted);
	if (!m) return /\d/.test(formatted) ? 0 : undefined;
	const [, separator, digits] = m;
	if (digits.length !== 3) return digits.length;
	const before = formatted.slice(0, m.index);
	if (before.includes(separator === "." ? "," : ".")) return 3;
	const whole = before.replace(/\D/g, "");
	const asDecimals = Number(`${whole || "0"}.${digits}`);
	const asThousands = Number(`${whole}${digits}`);
	const amount = Math.abs(raw);
	return Math.abs(asDecimals - amount) < Math.abs(asThousands - amount) ? 3 : 0;
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
	// The symbol and any space next to it. Signs and brackets of negative amounts aren't part of it,
	// wherever Dataverse's negative currency format puts them ("$-1.00", "1.00-$", "1.00 $-")
	const sign = /[-−()]/g;
	const prefix = (/^[^\d]+/.exec(text)?.[0] ?? "").replace(sign, "").replace(/^\s+/, "");
	const suffix = (/[^\d]+$/.exec(text)?.[0] ?? "").replace(sign, "").replace(/\s+$/, "");
	const literal = (s: string) => `"${s.replace(/"/g, "")}"`;
	if (prefix.trim()) return `${literal(prefix)}${number}`;
	if (suffix.trim()) return `${number}${literal(suffix)}`;
	return number;
}
