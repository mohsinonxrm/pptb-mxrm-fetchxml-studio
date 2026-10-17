/**
 * Field building blocks for the property editors (#90b), Fluent UI v9 only.
 * - InfoLabel instead of a Tooltip on an icon: Fluent says tooltips must not sit on static icons,
 *   and InfoLabel is keyboard accessible and can hold a "Learn more" link.
 * - Checkbox keeps its own label (Fluent: don't use Field's label for a Checkbox); Field is only
 *   added for a hint or validation message.
 * - SpinButton for bounded integers; empty means "not set".
 */
import type { ReactElement, ReactNode } from "react";
import { Button, Checkbox, Field, InfoLabel, Label, Link, SpinButton, Text, type LabelProps } from "@fluentui/react-components";
import { useEditorStyles } from "./editorStyles";

export type Content = string | ReactElement;
export type Validation = { state: "error" | "warning" | "none"; message: string } | undefined;

/** InfoLabel popover content: text plus an optional "Learn more" link to Microsoft Learn. */
export function Info({ text, href }: { text: ReactNode; href?: string }) {
	return (
		<>
			{text}
			{href && (
				<>
					{" "}
					<Link href={href} target="_blank" rel="noreferrer" inline>
						Learn more
					</Link>
				</>
			)}
		</>
	);
}

export function SectionHeading({ children }: { children: ReactNode }) {
	return <Label weight="semibold">{children}</Label>;
}

const infoLabel = (label: Content, info?: Content) =>
	info
		? {
				children: (_: unknown, props: LabelProps) => (
					<InfoLabel {...props} info={info}>
						{label}
					</InfoLabel>
				),
			}
		: label;

/** Field with an InfoLabel (Fluent "Field › Info button" pattern). */
export function InfoField({
	label,
	info,
	hint,
	required,
	validation,
	children,
}: {
	label: Content;
	info?: Content;
	hint?: Content;
	required?: boolean;
	validation?: Validation;
	children: ReactElement;
}) {
	return (
		<Field
			label={infoLabel(label, info)}
			hint={validation?.message ? undefined : hint}
			required={required}
			validationState={validation?.state}
			validationMessage={validation?.message}
		>
			{children}
		</Field>
	);
}

/** Checkbox whose own label is an InfoLabel; wrapped in a Field only for a hint or validation. */
export function InfoCheckbox({
	label,
	info,
	checked,
	onChange,
	disabled,
	hint,
	validation,
}: {
	label: Content;
	info?: Content;
	checked: boolean;
	onChange: (checked: boolean) => void;
	disabled?: boolean;
	hint?: Content;
	validation?: Validation;
}) {
	const box = (
		<Checkbox
			checked={checked}
			disabled={disabled}
			onChange={(_, data) => onChange(data.checked === true)}
			label={infoLabel(label, info)}
		/>
	);
	if (!hint && !validation?.message) return box;
	return (
		<Field
			hint={validation?.message ? undefined : hint}
			validationState={validation?.state}
			validationMessage={validation?.message}
		>
			{box}
		</Field>
	);
}

/** SpinButton field: empty = not set; out-of-range values are flagged, never clamped silently. */
export function NumberField({
	label,
	info,
	hint,
	value,
	min,
	max,
	rangeSeverity = "error",
	placeholder,
	disabled,
	onChange,
	validation,
}: {
	label: Content;
	info?: Content;
	hint?: Content;
	value: number | undefined;
	min?: number;
	max?: number;
	/** How to flag values outside min/max (warning when the limit is documented but not enforced) */
	rangeSeverity?: "error" | "warning";
	placeholder?: string;
	disabled?: boolean;
	onChange: (value: number | undefined) => void;
	validation?: Validation;
}) {
	const outOfRange =
		value !== undefined && ((min !== undefined && value < min) || (max !== undefined && value > max));
	const range: Validation = outOfRange
		? {
				state: rangeSeverity,
				message:
					max === undefined
						? `Must be at least ${min?.toLocaleString()}.`
						: `Must be from ${min?.toLocaleString()} to ${max.toLocaleString()}.`,
			}
		: undefined;
	return (
		<InfoField label={label} info={info} hint={hint} validation={validation ?? range}>
			<SpinButton
				value={value ?? null}
				min={min}
				max={max}
				placeholder={placeholder}
				disabled={disabled}
				onChange={(_, data) => {
					if (data.value !== undefined && data.value !== null) return onChange(data.value);
					const raw = (data.displayValue ?? "").trim();
					if (raw === "") return onChange(undefined);
					const parsed = Number(raw);
					if (Number.isFinite(parsed)) onChange(Math.trunc(parsed));
				}}
			/>
		</InfoField>
	);
}

/** A setting that doesn't apply where it is (parsed XML or after a mode change), with a Remove action. */
export function LeftoverSetting({
	text,
	validation,
	onRemove,
}: {
	text: string;
	validation: NonNullable<Validation>;
	onRemove: () => void;
}) {
	const styles = useEditorStyles();
	return (
		<Field validationState={validation.state} validationMessage={validation.message}>
			<div className={styles.leftover}>
				<Text>{text}</Text>
				<Button size="small" onClick={onRemove}>
					Remove
				</Button>
			</div>
		</Field>
	);
}
