import { describe, expect, it } from "vitest";
import {
	UTC_TIME_ZONE,
	dateValueContext,
	formatConditionDate,
	formatOffset,
	parseConditionDate,
	pickerDateToWallClock,
	resultWallClock,
	timeZoneFromUserSettings,
	toWallClock,
	utcOffsetMinutes,
	wallClockToPickerDate,
	wallClockToUtcMs,
	type DataverseTimeZone,
	type WallClock,
} from "./dateTimeValues";

/** UserSettings rows as Dataverse returns them */
const PACIFIC = timeZoneFromUserSettings(
	{
		timezonecode: 4,
		timezonebias: 480,
		timezonestandardbias: 0,
		timezonedaylightbias: -60,
		timezonedaylightmonth: 3,
		timezonedaylightday: 2,
		timezonedaylightdayofweek: 0,
		timezonedaylighthour: 2,
		timezonestandardmonth: 11,
		timezonestandardday: 1,
		timezonestandarddayofweek: 0,
		timezonestandardhour: 2,
	},
	"Pacific Standard Time",
);
const SYDNEY = timeZoneFromUserSettings(
	{
		timezonebias: -600,
		timezonedaylightbias: -60,
		timezonedaylightmonth: 10,
		timezonedaylightday: 1,
		timezonedaylightdayofweek: 0,
		timezonedaylighthour: 2,
		timezonestandardmonth: 4,
		timezonestandardday: 1,
		timezonestandarddayofweek: 0,
		timezonestandardhour: 3,
	},
	"AUS Eastern Standard Time",
);
const TOKYO: DataverseTimeZone = timeZoneFromUserSettings({ timezonebias: -540 }, "Tokyo Standard Time");
// The sandbox user's row while the time zone was UTC (every rule column is 0)
const UTC_ROW = timeZoneFromUserSettings({ timezonecode: 92, timezonebias: 0, timezonedaylightmonth: 0, timezonestandardmonth: 0 }, "UTC");

const wall = (year: number, month: number, day: number, hour = 0, minute = 0, second = 0): WallClock => ({ year, month, day, hour, minute, second });
const ms = (iso: string) => Date.parse(iso);

describe("the user's Dataverse time zone (UserSettings rules)", () => {
	it("reads offsets without daylight saving", () => {
		expect(utcOffsetMinutes(UTC_ROW, ms("2024-07-01T00:00:00Z"))).toBe(0);
		expect(utcOffsetMinutes(TOKYO, ms("2024-07-01T00:00:00Z"))).toBe(540);
		expect(UTC_ROW.daylight).toBeUndefined();
	});

	it("follows daylight saving in the northern hemisphere (Pacific: 2nd Sunday of March to 1st Sunday of November)", () => {
		expect(utcOffsetMinutes(PACIFIC, ms("2024-01-15T12:00:00Z"))).toBe(-480);
		expect(utcOffsetMinutes(PACIFIC, ms("2024-07-15T12:00:00Z"))).toBe(-420);
		// 2024-03-10 02:00 PST = 10:00Z
		expect(utcOffsetMinutes(PACIFIC, ms("2024-03-10T09:59:00Z"))).toBe(-480);
		expect(utcOffsetMinutes(PACIFIC, ms("2024-03-10T10:00:00Z"))).toBe(-420);
		// 2024-11-03 02:00 PDT = 09:00Z
		expect(utcOffsetMinutes(PACIFIC, ms("2024-11-03T08:59:00Z"))).toBe(-420);
		expect(utcOffsetMinutes(PACIFIC, ms("2024-11-03T09:00:00Z"))).toBe(-480);
	});

	it("follows daylight saving across the new year in the southern hemisphere (Sydney)", () => {
		expect(utcOffsetMinutes(SYDNEY, ms("2024-01-15T00:00:00Z"))).toBe(660);
		expect(utcOffsetMinutes(SYDNEY, ms("2024-07-15T00:00:00Z"))).toBe(600);
		expect(utcOffsetMinutes(SYDNEY, ms("2024-12-25T00:00:00Z"))).toBe(660);
	});

	it("converts between instants and wall clocks", () => {
		expect(toWallClock(ms("2024-03-14T09:00:00Z"), TOKYO)).toEqual(wall(2024, 3, 14, 18));
		expect(toWallClock(ms("2024-03-14T20:00:00Z"), TOKYO)).toEqual(wall(2024, 3, 15, 5));
		expect(new Date(wallClockToUtcMs(wall(2024, 7, 1, 9), PACIFIC)).toISOString()).toBe("2024-07-01T16:00:00.000Z");
		expect(new Date(wallClockToUtcMs(wall(2024, 1, 1, 9), PACIFIC)).toISOString()).toBe("2024-01-01T17:00:00.000Z");
	});

	it("formats offsets", () => {
		expect(formatOffset(540)).toBe("+09:00");
		expect(formatOffset(-420)).toBe("-07:00");
		expect(formatOffset(0)).toBe("+00:00");
		expect(formatOffset(330)).toBe("+05:30");
		expect(formatOffset(-210)).toBe("-03:30");
	});
});

describe("writing a picked date into a condition (#207)", () => {
	const dt = { Format: "DateAndTime", DateTimeBehavior: { Value: "UserLocal" } };
	const tzi = { Format: "DateAndTime", DateTimeBehavior: { Value: "TimeZoneIndependent" } };
	const dateOnly = { Format: "DateOnly", DateTimeBehavior: { Value: "DateOnly" } };
	const userLocalDateFormat = { Format: "DateOnly", DateTimeBehavior: { Value: "UserLocal" } };
	const picked = wall(2024, 3, 14, 9, 0);

	it("User Local: the wall clock with the user's Dataverse offset at that time", () => {
		expect(formatConditionDate(picked, dateValueContext(dt, "eq", TOKYO))).toBe("2024-03-14T09:00:00+09:00");
		expect(formatConditionDate(picked, dateValueContext(dt, "gt", UTC_TIME_ZONE))).toBe("2024-03-14T09:00:00+00:00");
		// March 14 is daylight time in Pacific, January isn't
		expect(formatConditionDate(picked, dateValueContext(dt, "lt", PACIFIC))).toBe("2024-03-14T09:00:00-07:00");
		expect(formatConditionDate(wall(2024, 1, 14, 9), dateValueContext(dt, "lt", PACIFIC))).toBe("2024-01-14T09:00:00-08:00");
	});

	it("User Local, time zone unknown: no offset (Dataverse reads it in the user's time zone, verified live)", () => {
		expect(formatConditionDate(picked, dateValueContext(dt, "eq", undefined))).toBe("2024-03-14T09:00:00");
	});

	it("Time Zone Independent: the wall clock, no offset", () => {
		expect(formatConditionDate(picked, dateValueContext(tzi, "eq", TOKYO))).toBe("2024-03-14T09:00:00");
	});

	it("Date Only columns, date-only formats and the on operators: YYYY-MM-DD", () => {
		expect(formatConditionDate(picked, dateValueContext(dateOnly, "eq", TOKYO))).toBe("2024-03-14");
		expect(formatConditionDate(picked, dateValueContext(userLocalDateFormat, "ge", TOKYO))).toBe("2024-03-14");
		for (const op of ["on", "on-or-before", "on-or-after"]) {
			expect(formatConditionDate(picked, dateValueContext(dt, op, SYDNEY))).toBe("2024-03-14");
		}
	});

	it("writes the same date text in every time zone (the issue's acceptance)", () => {
		for (const tz of [UTC_TIME_ZONE, TOKYO, PACIFIC, SYDNEY]) {
			expect(formatConditionDate(picked, dateValueContext(dt, "on", tz))).toBe("2024-03-14");
			expect(formatConditionDate(picked, dateValueContext(dateOnly, "eq", tz))).toBe("2024-03-14");
		}
	});
});

describe("showing a condition's value in the picker (#198, #207)", () => {
	const dt = { Format: "DateAndTime", DateTimeBehavior: { Value: "UserLocal" } };
	const tzi = { Format: "DateAndTime", DateTimeBehavior: { Value: "TimeZoneIndependent" } };
	const dateOnly = { Format: "DateOnly", DateTimeBehavior: { Value: "DateOnly" } };

	it("a YYYY-MM-DD value is that date, never the day before (#198)", () => {
		for (const tz of [UTC_TIME_ZONE, TOKYO, PACIFIC, SYDNEY]) {
			expect(parseConditionDate("2024-03-14", dateValueContext(dateOnly, "eq", tz))).toEqual(wall(2024, 3, 14));
			expect(parseConditionDate("2024-03-14", dateValueContext(dt, "on", tz))).toEqual(wall(2024, 3, 14));
		}
	});

	it("User Local: an instant is shown in the user's Dataverse time zone", () => {
		expect(parseConditionDate("2024-03-14T09:00:00Z", dateValueContext(dt, "eq", TOKYO))).toEqual(wall(2024, 3, 14, 18));
		expect(parseConditionDate("2024-03-14T02:00:00-07:00", dateValueContext(dt, "eq", UTC_TIME_ZONE))).toEqual(wall(2024, 3, 14, 9));
		expect(parseConditionDate("2024-03-13T23:00:00.000Z", dateValueContext(dt, "eq", SYDNEY))).toEqual(wall(2024, 3, 14, 10));
		// Without an offset, the value is already in the user's time zone
		expect(parseConditionDate("2024-03-14T18:00:00", dateValueContext(dt, "eq", TOKYO))).toEqual(wall(2024, 3, 14, 18));
	});

	it("the on operators: Z is converted, an explicit offset is ignored (as Dataverse does)", () => {
		expect(parseConditionDate("2024-03-14T20:00:00Z", dateValueContext(dt, "on", TOKYO))).toMatchObject({ year: 2024, month: 3, day: 15 });
		expect(parseConditionDate("2024-03-14T23:30:00-08:00", dateValueContext(dt, "on", UTC_TIME_ZONE))).toMatchObject({ day: 14 });
	});

	it("Date Only and Time Zone Independent: the value as written; Z and offsets are ignored", () => {
		expect(parseConditionDate("2024-03-14T23:30:00-08:00", dateValueContext(dateOnly, "eq", UTC_TIME_ZONE))).toMatchObject({ day: 14 });
		expect(parseConditionDate("2024-03-14T09:00:00+09:00", dateValueContext(tzi, "eq", PACIFIC))).toEqual(wall(2024, 3, 14, 9));
		expect(parseConditionDate("2024-03-14T09:00:00Z", dateValueContext(tzi, "eq", TOKYO))).toEqual(wall(2024, 3, 14, 9));
	});

	it("round-trips what it writes", () => {
		for (const tz of [UTC_TIME_ZONE, TOKYO, PACIFIC, SYDNEY]) {
			for (const [attr, op] of [[dt, "eq"], [tzi, "lt"], [dateOnly, "eq"], [dt, "on"]] as const) {
				const ctx = dateValueContext(attr, op, tz);
				const picked = wall(2024, 11, 3, 1, 30);
				const written = formatConditionDate(picked, ctx);
				const expected = ctx.dateOnly ? wall(2024, 11, 3) : picked;
				expect(parseConditionDate(written, ctx)).toEqual(expected);
			}
		}
	});

	it("leaves values it can't read alone", () => {
		const ctx = dateValueContext(dt, "eq", TOKYO);
		expect(parseConditionDate("@today", ctx)).toBeUndefined();
		expect(parseConditionDate("2024-02-30", ctx)).toBeUndefined();
		expect(parseConditionDate("", ctx)).toBeUndefined();
	});

	it("carries a wall clock through the picker's Date unchanged, whatever the browser's time zone", () => {
		const w = wall(2024, 3, 14, 9, 30, 15);
		expect(pickerDateToWallClock(wallClockToPickerDate(w))).toEqual(w);
	});
});

describe("result values for Excel (#184)", () => {
	it("User Local: in the user's Dataverse time zone", () => {
		expect(resultWallClock("2024-03-14T20:00:00Z", "UserLocal", TOKYO)).toEqual(wall(2024, 3, 15, 5));
		expect(resultWallClock("2024-07-01T16:00:00Z", "UserLocal", PACIFIC)).toEqual(wall(2024, 7, 1, 9));
	});

	it("Time Zone Independent: the stored reading", () => {
		expect(resultWallClock("2024-03-14T23:30:00Z", "TimeZoneIndependent", TOKYO)).toEqual(wall(2024, 3, 14, 23, 30));
	});

	it("Date Only: the date", () => {
		expect(resultWallClock("2024-03-14", "DateOnly", PACIFIC)).toEqual(wall(2024, 3, 14));
	});
});
