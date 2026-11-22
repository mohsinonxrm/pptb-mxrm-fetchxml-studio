/**
 * Power Automate on the Code tab (#233): what to type into the Microsoft Dataverse connector's "List
 * rows" action, field by field. Each field is labelled as the connector labels it, with its hint and
 * its own Copy; Copy all is in the language row (CodePanel).
 */
import {
	Button,
	Caption1,
	Field,
	Input,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Subtitle2,
	Textarea,
	Tooltip,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Copy20Regular, Flash20Filled } from "@fluentui/react-icons";
import { useCopy } from "../../../../shared/feedback/useCopy";
import type {
	PowerAutomateActionSpec,
	PowerAutomateField,
} from "../../engine/fetchxmlCodeGenerators";

const useStyles = makeStyles({
	root: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalL,
		maxWidth: "880px",
	},
	heading: {
		display: "flex",
		alignItems: "baseline",
		flexWrap: "wrap",
		columnGap: tokens.spacingHorizontalS,
	},
	headingIcon: { color: tokens.colorBrandForeground1, alignSelf: "center" },
	connector: { color: tokens.colorNeutralForeground3 },
	// Field (label, value, hint) | Copy, lined up with the value
	row: {
		display: "grid",
		gridTemplateColumns: "minmax(0, 1fr) auto",
		columnGap: tokens.spacingHorizontalS,
		alignItems: "start",
	},
	value: { fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase200 },
	// The Field's label line, so Copy sits beside the value rather than the label
	copy: { marginTop: "22px" },
	notes: { margin: `${tokens.spacingVerticalXS} 0 0`, paddingLeft: tokens.spacingHorizontalL },
});

export interface PowerAutomatePaneProps {
	spec: PowerAutomateActionSpec;
}

export function PowerAutomatePane({ spec }: PowerAutomatePaneProps) {
	const styles = useStyles();
	return (
		<div className={styles.root}>
			<div className={styles.heading}>
				<Flash20Filled className={styles.headingIcon} />
				<Subtitle2 as="h3">{spec.actionName}</Subtitle2>
				<Caption1 className={styles.connector}>
					{spec.connector} · labels match the connector
				</Caption1>
			</div>

			{spec.banner && (
				<MessageBar layout="multiline" intent="info">
					<MessageBarBody>{spec.banner}</MessageBarBody>
				</MessageBar>
			)}

			{spec.fields.map((f) => (
				<PowerAutomateRow key={f.label} field={f} />
			))}

			{spec.notes && spec.notes.length > 0 && (
				<MessageBar layout="multiline" intent="info">
					<MessageBarBody>
						<MessageBarTitle>Notes</MessageBarTitle>
						<ul className={styles.notes}>
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

function PowerAutomateRow({ field }: { field: PowerAutomateField }) {
	const styles = useStyles();
	const copy = useCopy();
	const empty = !field.value;
	const copyName = empty ? `Copy ${field.label}: it's empty` : `Copy ${field.label}`;
	return (
		<div className={styles.row}>
			<Field label={field.label} hint={field.hint}>
				{field.multiline ? (
					<Textarea
						// The font goes on the <textarea> itself, which sets its own
						textarea={{ className: styles.value }}
						value={field.value}
						readOnly
						resize="vertical"
						rows={Math.min(6, Math.max(2, field.value.split("\n").length))}
					/>
				) : (
					<Input className={styles.value} value={field.value} readOnly />
				)}
			</Field>
			<Tooltip content={copyName} relationship="label">
				<Button
					className={styles.copy}
					appearance="subtle"
					icon={<Copy20Regular />}
					disabledFocusable={empty}
					// "Copied Fetch Xml Query" or "Couldn't copy to the clipboard" (#277)
					onClick={empty ? undefined : () => void copy(field.value, field.label)}
				/>
			</Tooltip>
		</div>
	);
}
