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
	pickerDateToWallClock,
	timeZoneHint,
	wallClockToPickerDate,
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
	const [selectedDate, setSelectedDate] = useState<Date | null>(null);
	const [selectedTime, setSelectedTime] = useState<Date | null>(null);
	const [timePickerValue, setTimePickerValue] = useState<string>("");

	// Sync selected date/time with prop value
	const { behavior, dateOnly } = ctx;
	useEffect(() => {
		const wall = value ? parseConditionDate(value, { behavior, dateOnly, timeZone }) : undefined;
		if (wall) {
			const dateValue = wallClockToPickerDate(wall);
			setSelectedDate(dateValue);
			if (!dateOnly) {
				setSelectedTime(dateValue);
				setTimePickerValue(formatDateToTimeString(dateValue));
			}
		} else {
			setSelectedDate(null);
			setSelectedTime(null);
			setTimePickerValue("");
		}
	}, [value, behavior, dateOnly, timeZone]);

	// Parse min/max dates
	const minDate = attribute.MinSupportedValue ? new Date(attribute.MinSupportedValue) : undefined;
	const maxDate = attribute.MaxSupportedValue ? new Date(attribute.MaxSupportedValue) : undefined;

	/** Writes the picked date with a time (or midnight) */
	const write = (date: Date, time: Date | null) => {
		const wall = pickerDateToWallClock(date);
		const withTime =
			!ctx.dateOnly && time
				? { ...wall, hour: time.getHours(), minute: time.getMinutes(), second: time.getSeconds() }
				: { ...wall, hour: 0, minute: 0, second: 0 };
		onChange(formatConditionDate(withTime, ctx));
	};

	// Handle date selection
	const handleDateChange = (date: Date | null | undefined) => {
		setSelectedDate(date ?? null);
		if (!date) {
			onChange(undefined);
			return;
		}
		write(date, selectedTime);
	};

	// Handle time selection (only when the time counts)
	const handleTimeChange = (
		_ev: unknown,
		data: { selectedTime: Date | null; selectedTimeText?: string }
	) => {
		setSelectedTime(data.selectedTime);
		setTimePickerValue(data.selectedTimeText ?? "");
		if (selectedDate && data.selectedTime) write(selectedDate, data.selectedTime);
	};

	// Handle time input (freeform)
	const handleTimeInput = (ev: React.ChangeEvent<HTMLInputElement>) => {
		setTimePickerValue(ev.target.value);
	};

	if (ctx.dateOnly) {
		// Only the date counts: just show DatePicker
		return (
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
					dateAnchor={selectedDate ?? undefined}
					selectedTime={selectedTime}
					onTimeChange={handleTimeChange}
					value={timePickerValue}
					onInput={handleTimeInput}
					disabled={disabled || !selectedDate}
					className={styles.timePicker}
				/>
			</div>
			{hint && <span className={styles.hint}>{hint}</span>}
		</div>
	);
}
