import {
	Input,
	Textarea,
	Label,
	Button,
	Tooltip,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Subtitle1,
	Subtitle2,
	Caption1,
	Badge,
	mergeClasses,
	tokens,
	makeStyles,
} from "@fluentui/react-components";
import { Copy20Regular, Flash20Filled } from "@fluentui/react-icons";
import { useCopy } from "../../../../shared/feedback/useCopy";
import type {
	PowerAutomateActionSpec,
	PowerAutomateField,
} from "../../engine/fetchxmlCodeGenerators";

// Form-style view that mirrors the Dataverse connector's "List rows" action UI.
// Each row is: label | read-only Input (or Textarea) | Copy button.

const usePAStyles = makeStyles({
	row: {
		display: "grid",
		gridTemplateColumns: "180px minmax(0, 1fr) auto",
		gap: "12px",
		alignItems: "start",
		marginBottom: "14px",
	},
	rowLabel: {
		paddingTop: "6px",
		fontWeight: 500,
		color: tokens.colorNeutralForeground1,
	},
	rowLabelEmpty: {
		color: tokens.colorNeutralForeground3,
	},
	hint: {
		fontSize: "10px",
		color: tokens.colorNeutralForeground3,
		marginTop: "4px",
		display: "block",
	},
	copyBtn: {
		minWidth: "90px",
	},
});

export interface PowerAutomatePaneProps {
	spec: PowerAutomateActionSpec;
}

export function PowerAutomatePane({ spec }: PowerAutomatePaneProps) {
	const copy = useCopy();
	const s = usePAStyles();
	const populatedCount = spec.fields.filter((f) => f.value).length;
	const allAsText = spec.fields
		.filter((f) => f.value)
		.map((f) => `${f.label.padEnd(28)} ${f.value}`)
		.join("\n");

	return (
		<div>
			{/* Header — action name + connector context */}
			<div
				style={{
					marginBottom: 16,
					display: "flex",
					alignItems: "baseline",
					gap: 12,
					flexWrap: "wrap",
				}}
			>
				<Flash20Filled style={{ color: tokens.colorBrandForeground1, marginTop: 6 }} />
				<Subtitle1 style={{ fontFamily: tokens.fontFamilyBase }}>
					{spec.actionName}
				</Subtitle1>
				<Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
					· {spec.connector}
				</Caption1>
				<span style={{ flexGrow: 1 }} />
				<Badge appearance="ghost">
					{populatedCount} of {spec.fields.length} populated
				</Badge>
				<Tooltip
					content="Copy every populated field as one block"
					relationship="description"
				>
					<Button
						size="small"
						appearance="outline"
						icon={<Copy20Regular />}
						onClick={() => void copy(allAsText, "all fields")}
					>
						Copy all
					</Button>
				</Tooltip>
			</div>

			{spec.banner && (
				<MessageBar layout="multiline" intent="info" style={{ marginBottom: 14 }}>
					<MessageBarBody>{spec.banner}</MessageBarBody>
				</MessageBar>
			)}

			{/* Section label */}
			<Subtitle2
				style={{
					display: "block",
					marginBottom: 10,
					color: tokens.colorNeutralForeground2,
				}}
			>
				{spec.actionName}
			</Subtitle2>

			<div>
				{spec.fields.map((f) => (
					<PARow key={f.label} field={f} styles={s} />
				))}
			</div>

			{/* Notes */}
			{spec.notes && spec.notes.length > 0 && (
				<MessageBar layout="multiline" intent="info" style={{ marginTop: 14 }}>
					<MessageBarBody>
						<MessageBarTitle>Notes</MessageBarTitle>
						<ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
							{spec.notes.map((n, i) => (
								<li key={i}>{n}</li>
							))}
						</ul>
					</MessageBarBody>
				</MessageBar>
			)}
		</div>
	);
}

function PARow({
	field,
	styles,
}: {
	field: PowerAutomateField;
	styles: ReturnType<typeof usePAStyles>;
}) {
	const copy = useCopy();

	// "Copied {field}" or "Couldn't copy to the clipboard" (#277)
	const onCopy = () => {
		if (!field.value) return;
		void copy(field.value, field.label);
	};

	return (
		<div className={styles.row}>
			<Label
				className={mergeClasses(
					styles.rowLabel,
					!field.value ? styles.rowLabelEmpty : undefined,
				)}
			>
				{field.label}
			</Label>

			<div style={{ display: "flex", flexDirection: "column", minWidth: 0 }}>
				{field.multiline ? (
					<Textarea
						value={field.value}
						readOnly
						rows={Math.min(4, Math.max(1, Math.ceil(field.value.length / 80)))}
						appearance="filled-lighter"
						style={{
							width: "100%",
							fontFamily: tokens.fontFamilyMonospace,
							fontSize: 12,
						}}
					/>
				) : (
					<Input
						value={field.value}
						readOnly
						appearance="filled-lighter"
						style={{
							width: "100%",
							fontFamily: tokens.fontFamilyMonospace,
							fontSize: 12,
						}}
					/>
				)}
				{field.hint && <Caption1 className={styles.hint}>{field.hint}</Caption1>}
			</div>

			<Tooltip
				content={field.value ? "Copy this value" : "Empty — nothing to copy"}
				relationship="label"
			>
				<Button
					size="medium"
					appearance={field.value ? "outline" : "subtle"}
					disabled={!field.value}
					icon={<Copy20Regular />}
					onClick={onCopy}
					className={styles.copyBtn}
				>
					Copy
				</Button>
			</Tooltip>
		</div>
	);
}
