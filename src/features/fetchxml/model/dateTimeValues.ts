/**
 * Date and time values in conditions (#207, #198) and the user's Dataverse time zone (#184)
 *
 * How Dataverse reads a condition value, verified live on a sandbox (2026-11-01), with the user's
 * Dataverse time zone set to UTC and then to Tokyo:
 * - on / on-or-before / on-or-after compare the record's date in the user's Dataverse time zone. A
 *   value ending in Z is converted to that time zone first; an explicit offset (+09:00) is ignored.
 * - Date Only columns: only the date counts; a time or offset is ignored.
 * - User Local columns, other operators: a value without an offset is read in the user's Dataverse
 *   time zone; an explicit offset is honoured.
 * - Time Zone Independent columns: compared as a wall clock; Z and offsets are ignored.
 *
 * So the tool writes `YYYY-MM-DD` for Date Only columns, date-only formats and the "on" operators,
 * the wall clock for Time Zone Independent columns, and the wall clock with the user's offset for
 * User Local columns ("2024-03-14T09:00:00+09:00"): it reads like what was picked and is exact.
 * https://learn.microsoft.com/power-apps/developer/data-platform/behavior-format-date-time-attribute
 */

/** A date and time as written on a clock, with no time zone */
export interface WallClock {
	year: number;
	/** 1–12 */
	month: number;
	day: number;
	hour: number;
	minute: number;
	second: number;
}

export type DateTimeBehaviorName = "UserLocal" | "DateOnly" | "TimeZoneIndependent";

/** DateTimeAttributeMetadata's behaviour; User Local when it isn't known (the default) */
export function dateTimeBehavior(attribute: { DateTimeBehavior?: { Value?: string } } | undefined): DateTimeBehaviorName {
	const value = attribute?.DateTimeBehavior?.Value;
	return value === "DateOnly" || value === "TimeZoneIndependent" ? value : "UserLocal";
}

/** Operators that compare dates, not times: their value is written as YYYY-MM-DD */
export const DATE_ONLY_OPERATORS: ReadonlySet<string> = new Set(["on", "on-or-before", "on-or-after"]);

// ─────────────────────────────────────────────────────────────────────────────
// The user's Dataverse time zone
// ─────────────────────────────────────────────────────────────────────────────

/** A daylight-saving transition, as UserSettings stores it (Windows TIME_ZONE_INFORMATION) */
export interface TransitionRule {
	/** 1–12 */
	month: number;
	/** Week of the month, 1–5; 5 is the last */
	week: number;
	/** 0 = Sunday */
	dayOfWeek: number;
	hour: number;
	minute: number;
}

/**
 * The user's Dataverse time zone (Personal Options), from UserSettings. Bias is in minutes with
 * UTC = local + bias, as in Windows; the daylight rules are absent where there's no daylight saving.
 */
export interface DataverseTimeZone {
	/** e.g. "UTC", "Tokyo Standard Time" */
	name: string;
	bias: number;
	standardBias: number;
	daylightBias: number;
	daylight?: TransitionRule;
	standard?: TransitionRule;
}

export const UTC_TIME_ZONE: DataverseTimeZone = { name: "UTC", bias: 0, standardBias: 0, daylightBias: 0 };

/** UserSettings columns for the time zone (select these) */
export const USER_SETTINGS_TIME_ZONE_COLUMNS = [
	"timezonecode",
	"timezonebias",
	"timezonestandardbias",
	"timezonedaylightbias",
	"timezonestandardmonth",
	"timezonestandardday",
	"timezonestandarddayofweek",
	"timezonestandardhour",
	"timezonestandardminute",
	"timezonedaylightmonth",
	"timezonedaylightday",
	"timezonedaylightdayofweek",
	"timezonedaylighthour",
	"timezonedaylightminute",
] as const;

/** The time zone from a UserSettings row and the time zone definition's standard name */
export function timeZoneFromUserSettings(settings: Record<string, unknown>, name: string): DataverseTimeZone {
	const n = (key: string) => (typeof settings[key] === "number" ? (settings[key] as number) : 0);
	const rule = (prefix: "timezonestandard" | "timezonedaylight"): TransitionRule | undefined =>
		n(`${prefix}month`) > 0
			? {
					month: n(`${prefix}month`),
					week: n(`${prefix}day`),
					dayOfWeek: n(`${prefix}dayofweek`),
					hour: n(`${prefix}hour`),
					minute: n(`${prefix}minute`),
				}
			: undefined;
	const daylight = rule("timezonedaylight");
	const standard = rule("timezonestandard");
	return {
		name,
		bias: n("timezonebias"),
		standardBias: n("timezonestandardbias"),
		daylightBias: n("timezonedaylightbias"),
		...(daylight && standard ? { daylight, standard } : {}),
	};
}

/** When a transition happens in a year, as the wall clock's milliseconds (Date.UTC of the local time) */
function transitionWallMs(year: number, rule: TransitionRule): number {
	const firstWeekday = new Date(Date.UTC(year, rule.month - 1, 1)).getUTCDay();
	let day = 1 + ((rule.dayOfWeek - firstWeekday + 7) % 7) + (rule.week - 1) * 7;
	const daysInMonth = new Date(Date.UTC(year, rule.month, 0)).getUTCDate();
	while (day > daysInMonth) day -= 7;
	return Date.UTC(year, rule.month - 1, day, rule.hour, rule.minute);
}

/** Minutes to add to UTC for the local time at an instant */
export function utcOffsetMinutes(tz: DataverseTimeZone, utcMs: number): number {
	if (!tz.daylight || !tz.standard) return 0 - (tz.bias + tz.standardBias);
	const year = new Date(utcMs).getUTCFullYear();
	// Daylight time starts at a standard-time clock reading and ends at a daylight-time one
	const daylightStarts = transitionWallMs(year, tz.daylight) + (tz.bias + tz.standardBias) * 60_000;
	const daylightEnds = transitionWallMs(year, tz.standard) + (tz.bias + tz.daylightBias) * 60_000;
	const inDaylight =
		daylightStarts < daylightEnds
			? utcMs >= daylightStarts && utcMs < daylightEnds
			: // Southern hemisphere: daylight time spans the new year
				utcMs >= daylightStarts || utcMs < daylightEnds;
	return 0 - (tz.bias + (inDaylight ? tz.daylightBias : tz.standardBias));
}

/** The wall clock in a time zone at an instant */
export function toWallClock(utcMs: number, tz: DataverseTimeZone): WallClock {
	return wallClockOf(new Date(utcMs + utcOffsetMinutes(tz, utcMs) * 60_000));
}

/** The instant a wall clock in a time zone stands for */
export function wallClockToUtcMs(wall: WallClock, tz: DataverseTimeZone): number {
	const asUtc = wallMs(wall);
	const first = asUtc - utcOffsetMinutes(tz, asUtc) * 60_000;
	return asUtc - utcOffsetMinutes(tz, first) * 60_000;
}

// ─────────────────────────────────────────────────────────────────────────────
// Condition values
// ─────────────────────────────────────────────────────────────────────────────

const wallMs = (w: WallClock) => Date.UTC(w.year, w.month - 1, w.day, w.hour, w.minute, w.second);
const wallClockOf = (d: Date): WallClock => ({
	year: d.getUTCFullYear(),
	month: d.getUTCMonth() + 1,
	day: d.getUTCDate(),
	hour: d.getUTCHours(),
	minute: d.getUTCMinutes(),
	second: d.getUTCSeconds(),
});
const pad = (n: number, width = 2) => String(Math.abs(n)).padStart(width, "0");

const formatDate = (w: WallClock) => `${pad(w.year, 4)}-${pad(w.month)}-${pad(w.day)}`;
const formatWall = (w: WallClock) => `${formatDate(w)}T${pad(w.hour)}:${pad(w.minute)}:${pad(w.second)}`;

/** "+09:00", "-07:00", "+00:00" */
export function formatOffset(minutes: number): string {
	return `${minutes < 0 ? "-" : "+"}${pad(Math.trunc(Math.abs(minutes) / 60))}:${pad(Math.abs(minutes) % 60)}`;
}

/** How a value is written and read: the column's behaviour, whether only the date counts, the time zone */
export interface DateValueContext {
	behavior: DateTimeBehaviorName;
	/** A date-only format, or an "on" operator */
	dateOnly: boolean;
	/** The user's Dataverse time zone; when unknown, User Local values are written without an offset (Dataverse then reads them in the user's time zone) */
	timeZone?: DataverseTimeZone;
}

/** The context for a condition on an attribute with an operator */
export function dateValueContext(
	attribute: { Format?: string; DateTimeBehavior?: { Value?: string } } | undefined,
	operator: string | undefined,
	timeZone: DataverseTimeZone | undefined,
): DateValueContext {
	const behavior = dateTimeBehavior(attribute);
	return {
		behavior,
		dateOnly: behavior === "DateOnly" || attribute?.Format === "DateOnly" || DATE_ONLY_OPERATORS.has(operator ?? ""),
		timeZone,
	};
}

/** A picked date and time, as the condition's value */
export function formatConditionDate(wall: WallClock, ctx: DateValueContext): string {
	if (ctx.dateOnly) return formatDate(wall);
	if (ctx.behavior === "TimeZoneIndependent" || !ctx.timeZone) return formatWall(wall);
	const offset = utcOffsetMinutes(ctx.timeZone, wallClockToUtcMs(wall, ctx.timeZone));
	return `${formatWall(wall)}${formatOffset(offset)}`;
}

const VALUE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*(Z|[+-]\d{2}:?\d{2})?)?$/i;

/**
 * A condition value as the wall clock to show in the picker, read the way Dataverse reads it
 * (see the top of this file); undefined when it isn't a date the picker can show. `YYYY-MM-DD`
 * is that date, never shifted by a time zone (#198).
 */
export function parseConditionDate(value: string, ctx: DateValueContext): WallClock | undefined {
	const m = VALUE_PATTERN.exec(value.trim());
	if (!m) return undefined;
	const [, y, mo, d, h = "0", mi = "0", s = "0", zone] = m;
	const literal: WallClock = { year: +y, month: +mo, day: +d, hour: +h, minute: +mi, second: +s };
	if (!isValidWall(literal)) return undefined;
	// Dataverse ignores the zone for Date Only and Time Zone Independent columns, and an explicit
	// offset (but not Z) for the "on" operators
	const zoneIgnored =
		!zone || ctx.behavior !== "UserLocal" || (ctx.dateOnly && zone.toUpperCase() !== "Z") || !ctx.timeZone;
	if (zoneIgnored) return literal;
	const offsetMinutes = zone.toUpperCase() === "Z" ? 0 : zoneMinutes(zone);
	return toWallClock(wallMs(literal) - offsetMinutes * 60_000, ctx.timeZone!);
}

function zoneMinutes(zone: string): number {
	const m = /^([+-])(\d{2}):?(\d{2})$/.exec(zone)!;
	return (m[1] === "-" ? -1 : 1) * (Number(m[2]) * 60 + Number(m[3]));
}

function isValidWall(w: WallClock): boolean {
	const d = new Date(wallMs(w));
	return d.getUTCFullYear() === w.year && d.getUTCMonth() + 1 === w.month && d.getUTCDate() === w.day && w.hour < 24 && w.minute < 60 && w.second < 60;
}

/** Which time zone a picked time is in, for the hint under the time picker (#207) */
export function timeZoneHint(ctx: DateValueContext, timeZone: DataverseTimeZone | undefined): string | undefined {
	if (ctx.dateOnly) return undefined;
	if (ctx.behavior === "TimeZoneIndependent") return "Time zone independent: no time zone is applied.";
	return timeZone?.name ? `Time in your Dataverse time zone (${timeZone.name}).` : "Time in your Dataverse time zone.";
}

// ─────────────────────────────────────────────────────────────────────────────
// Pickers and Excel
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Fluent's date and time pickers work with Date objects in the browser's time zone; a wall clock
 * is carried as the Date with the same local reading, whatever the browser's time zone.
 */
export const wallClockToPickerDate = (w: WallClock): Date => new Date(w.year, w.month - 1, w.day, w.hour, w.minute, w.second);

export const pickerDateToWallClock = (d: Date): WallClock => ({
	year: d.getFullYear(),
	month: d.getMonth() + 1,
	day: d.getDate(),
	hour: d.getHours(),
	minute: d.getMinutes(),
	second: d.getSeconds(),
});

/**
 * A result value (the Web API's raw value) as the wall clock to show (#184): Date Only values are
 * their date, Time Zone Independent values are their stored reading, User Local values are the
 * instant in the user's Dataverse time zone.
 */
export function resultWallClock(raw: string, behavior: DateTimeBehaviorName, timeZone: DataverseTimeZone): WallClock | undefined {
	const m = VALUE_PATTERN.exec(raw.trim());
	if (!m) return undefined;
	const [, y, mo, d, h = "0", mi = "0", s = "0", zone] = m;
	const literal: WallClock = { year: +y, month: +mo, day: +d, hour: +h, minute: +mi, second: +s };
	if (!isValidWall(literal)) return undefined;
	if (behavior === "DateOnly") return { ...literal, hour: 0, minute: 0, second: 0 };
	if (behavior === "TimeZoneIndependent" || !zone) return literal;
	const offsetMinutes = zone.toUpperCase() === "Z" ? 0 : zoneMinutes(zone);
	return toWallClock(wallMs(literal) - offsetMinutes * 60_000, timeZone);
}

/**
 * A wall clock as the Date to give exceljs, which writes a Date's UTC reading into the cell: the
 * cell then shows the wall clock, whatever the browser's time zone.
 */
export const wallClockToExcelDate = (w: WallClock): Date => new Date(wallMs(w));
