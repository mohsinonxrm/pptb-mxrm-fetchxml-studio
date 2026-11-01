/**
 * Shared DateTime Input Picker component using Fluent UI v9 DatePicker and TimePicker
 * Handles DateTime attribute types
 * Renders a DatePicker when only the date counts (Date Only columns and formats, and the "on"
 * operators), DatePicker + TimePicker otherwise.
 *
 * The value is the condition's text, written the way Dataverse reads it (#207, dateTimeValues.ts):
 * YYYY-MM-DD when only the date counts, the wall clock for Time Zone Independent columns, and the
 * wall clock with the user's Dataverse offset for User Local columns. The pickers show that wall
 * clock, whatever the browser's time zone (#198). A value is only rewritten when a date or time is
 * picked, so loaded FetchXML keeps its values as written.
 */

import { useState, useEffect } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { TimePicker, formatDateToTimeString } from "@fluentui/react-timepicker-compat";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";
import {
	dateValueContext,
	formatConditionDate,
	parseConditionDate,
	pickerDate,
	pickerTime,
	timeZoneHint,
	withPickedDate,
	withPickedTime,
	TIME_PICKER_ANCHOR,
	type WallClock,
} from "../../features/fetchxml/model/dateTimeValues";
import { useDataverseTimeZone } from "../hooks/useDataverseTimeZone";

const useStyles = makeStyles({
	root: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalXXS,
	},
	pickers: {
		display: "flex",
		flexDirection: "row",
		gap: "8px",
	},
	datePicker: {
		flexGrow: 1,
	},
	timePicker: {
		flexGrow: 1,
	},
	// Like a Field hint
	hint: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
		lineHeight: tokens.lineHeightBase200,
	},
});

interface DateTimeInputPickerProps {
	attribute: AttributeMetadata;
	/** The condition's value as written */
	value?: string;
	/** The condition's operator: the "on" operators take a date only */
	operator?: string;
	onChange: (value: string | undefined) => void;
	placeholder?: string;
	disabled?: boolean;
	/** Show which time zone the time is in (the between picker shows it once, under both) */
	showTimeZoneHint?: boolean;
}

export function DateTimeInputPicker({
	attribute,
	value,
	operator,
	onChange,
	placeholder = "Select a date",
	disabled = false,
	showTimeZoneHint = true,
}: DateTimeInputPickerProps) {
	const styles = useStyles();
	const timeZone = useDataverseTimeZone();
	const ctx = dateValueContext(attribute, operator, timeZone);
	// The value's wall clock: the pickers get Dates made from it, never the other way round (#207)
	const [wall, setWall] = useState<WallClock | undefined>(undefined);
	const [timePickerValue, setTimePickerValue] = useState<string>("");

	// Sync the wall clock with the value; nothing is written back here, so loaded values stay as written
	const { behavior, dateOnly, dateOperator } = ctx;
	useEffect(() => {
		const parsed = value ? parseConditionDate(value, { behavior, dateOnly, dateOperator, timeZone }) : undefined;
		setWall(parsed);
		setTimePickerValue(parsed && !dateOnly ? formatDateToTimeString(pickerTime(parsed)) : "");
	}, [value, behavior, dateOnly, dateOperator, timeZone]);

	const selectedDate = wall ? pickerDate(wall) : null;
	const selectedTime = wall && !ctx.dateOnly ? pickerTime(wall) : null;
	// A value the picker can't show (or an instant while the time zone is unknown) is kept as it is
	const unreadable = value && !wall ? `Current value: ${value}` : undefined;

	// Parse min/max dates
	const minDate = attribute.MinSupportedValue ? new Date(attribute.MinSupportedValue) : undefined;
	const maxDate = attribute.MaxSupportedValue ? new Date(attribute.MaxSupportedValue) : undefined;

	/** Writes a picked wall clock */
	const write = (next: WallClock) => {
		setWall(next);
		onChange(formatConditionDate(next, ctx));
	};

	// Handle date selection
	const handleDateChange = (date: Date | null | undefined) => {
		if (!date) {
			setWall(undefined);
			onChange(undefined);
			return;
		}
		write(withPickedDate(wall, date, ctx.dateOnly));
	};

	// Handle time selection (only when the time counts)
	const handleTimeChange = (
		_ev: unknown,
		data: { selectedTime: Date | null; selectedTimeText?: string }
	) => {
		setTimePickerValue(data.selectedTimeText ?? "");
		if (wall && data.selectedTime) write(withPickedTime(wall, data.selectedTime));
	};

	// Handle time input (freeform)
	const handleTimeInput = (ev: React.ChangeEvent<HTMLInputElement>) => {
		setTimePickerValue(ev.target.value);
	};

	if (ctx.dateOnly) {
		// Only the date counts: just show DatePicker
		return (
			<div className={styles.root}>
				<DatePicker
					value={selectedDate}
					onSelectDate={handleDateChange}
					placeholder={placeholder}
					disabled={disabled}
					minDate={minDate}
					maxDate={maxDate}
					allowTextInput
					className={styles.datePicker}
				/>
				{unreadable && <span className={styles.hint}>{unreadable}</span>}
			</div>
		);
	}

	const hint = showTimeZoneHint ? timeZoneHint(ctx, timeZone) : undefined;
	// Date and time: show DatePicker + TimePicker
	return (
		<div className={styles.root}>
			<div className={styles.pickers}>
				<DatePicker
					value={selectedDate}
					onSelectDate={handleDateChange}
					placeholder="Select date"
					disabled={disabled}
					minDate={minDate}
					maxDate={maxDate}
					allowTextInput
					className={styles.datePicker}
				/>
				<TimePicker
					placeholder="Select time"
					freeform
					dateAnchor={TIME_PICKER_ANCHOR}
					selectedTime={selectedTime}
					onTimeChange={handleTimeChange}
					value={timePickerValue}
					onInput={handleTimeInput}
					disabled={disabled || !selectedDate}
					className={styles.timePicker}
				/>
			</div>
			{unreadable && <span className={styles.hint}>{unreadable}</span>}
			{hint && <span className={styles.hint}>{hint}</span>}
		</div>
	);
}
