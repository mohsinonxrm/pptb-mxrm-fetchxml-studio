/**
 * Shared DateTime Input Picker component using Fluent UI v9 DatePicker and TimePicker
 * Handles DateTime attribute types
 * Renders DatePicker for DateOnly format, DatePicker + TimePicker for DateAndTime format
 */

import { useState, useEffect } from "react";
import { makeStyles } from "@fluentui/react-components";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { TimePicker, formatDateToTimeString } from "@fluentui/react-timepicker-compat";
import type { AttributeMetadata } from "../../features/fetchxml/api/pptbClient";

const useStyles = makeStyles({
	root: {
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
});

interface DateTimeInputPickerProps {
	attribute: AttributeMetadata;
	value?: Date | string;
	onChange: (value: Date | undefined) => void;
	placeholder?: string;
	disabled?: boolean;
}

export function DateTimeInputPicker({
	attribute,
	value,
	onChange,
	placeholder = "Select a date",
	disabled = false,
}: DateTimeInputPickerProps) {
	const styles = useStyles();
	const [selectedDate, setSelectedDate] = useState<Date | null>(null);
	const [selectedTime, setSelectedTime] = useState<Date | null>(null);
	const [timePickerValue, setTimePickerValue] = useState<string>("");

	const isDateOnly = attribute.Format === "DateOnly";

	// Sync selected date/time with prop value
	useEffect(() => {
		if (value) {
			const dateValue = typeof value === "string" ? new Date(value) : value;
			if (!isNaN(dateValue.getTime())) {
				setSelectedDate(dateValue);
				if (!isDateOnly) {
					setSelectedTime(dateValue);
					setTimePickerValue(formatDateToTimeString(dateValue));
				}
			}
		} else {
			setSelectedDate(null);
			setSelectedTime(null);
			setTimePickerValue("");
		}
	}, [value, isDateOnly]);

	// Parse min/max dates
	const minDate = attribute.MinSupportedValue ? new Date(attribute.MinSupportedValue) : undefined;
	const maxDate = attribute.MaxSupportedValue ? new Date(attribute.MaxSupportedValue) : undefined;

	// Handle date selection
	const handleDateChange = (date: Date | null | undefined) => {
		setSelectedDate(date ?? null);

		if (!date) {
			onChange(undefined);
			return;
		}

		if (isDateOnly) {
			// For DateOnly, just use the date at midnight UTC
			const dateOnly = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
			onChange(dateOnly);
		} else {
			// For DateAndTime, combine with selected time or use midnight
			if (selectedTime) {
				const combined = new Date(
					date.getFullYear(),
					date.getMonth(),
					date.getDate(),
					selectedTime.getHours(),
					selectedTime.getMinutes(),
					selectedTime.getSeconds()
				);
				onChange(combined);
			} else {
				onChange(date);
			}
		}
	};

	// Handle time selection (only for DateAndTime format)
	const handleTimeChange = (
		_ev: unknown,
		data: { selectedTime: Date | null; selectedTimeText?: string }
	) => {
		setSelectedTime(data.selectedTime);
		setTimePickerValue(data.selectedTimeText ?? "");

		if (selectedDate && data.selectedTime) {
			const combined = new Date(
				selectedDate.getFullYear(),
				selectedDate.getMonth(),
				selectedDate.getDate(),
				data.selectedTime.getHours(),
				data.selectedTime.getMinutes(),
				data.selectedTime.getSeconds()
			);
			onChange(combined);
		}
	};

	// Handle time input (freeform)
	const handleTimeInput = (ev: React.ChangeEvent<HTMLInputElement>) => {
		setTimePickerValue(ev.target.value);
	};

	if (isDateOnly) {
		// DateOnly: just show DatePicker
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

	// DateAndTime: show DatePicker + TimePicker
	return (
		<div className={styles.root}>
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
	);
}
