/**
 * The Keyboard shortcuts dialog (#186): Ctrl+/ or More › Keyboard shortcuts. Built from the shortcut
 * table (`src/app/shortcuts/registry.ts`), so it always matches what the keys do: the keys that work
 * anywhere in the tool first, then the ones handled where focus is. Keys that are planned but not
 * handled yet (`pending`) aren't listed. A description list of labels
 * and keys (design D's ShortcutList), with each key in a <kbd>; on a Mac, Ctrl shows as ⌘.
 */
import {
	Body1,
	Button,
	Dialog,
	DialogActions,
	DialogBody,
	DialogContent,
	DialogSurface,
	DialogTitle,
	Subtitle2,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Fragment } from "react";
import { SHORTCUTS, formatKeys, type ShortcutScope } from "../../../../app/shortcuts/registry";

const useStyles = makeStyles({
	section: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
		marginBlockEnd: tokens.spacingVerticalL,
	},
	list: {
		display: "grid",
		gridTemplateColumns: "1fr auto",
		columnGap: tokens.spacingHorizontalL,
		rowGap: tokens.spacingVerticalS,
		margin: 0,
	},
	term: { margin: 0, minWidth: 0 },
	keys: {
		margin: 0,
		display: "flex",
		flexWrap: "wrap",
		alignItems: "center",
		justifyContent: "end",
		gap: tokens.spacingHorizontalXS,
		color: tokens.colorNeutralForeground3,
	},
	kbd: {
		fontFamily: tokens.fontFamilyBase,
		fontSize: tokens.fontSizeBase200,
		lineHeight: tokens.lineHeightBase200,
		color: tokens.colorNeutralForeground1,
		backgroundColor: tokens.colorNeutralBackground3,
		border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke1}`,
		borderRadius: tokens.borderRadiusSmall,
		paddingInline: tokens.spacingHorizontalXS,
		whiteSpace: "nowrap",
		"@media (forced-colors: active)": { border: `${tokens.strokeWidthThin} solid CanvasText` },
	},
});

const SECTIONS: { scope: ShortcutScope; title: string }[] = [
	{ scope: "global", title: "Anywhere in the tool" },
	{ scope: "documented", title: "Where focus is" },
];

export interface ShortcutsDialogProps {
	open: boolean;
	onClose: () => void;
}

export function ShortcutsDialog({ open, onClose }: ShortcutsDialogProps) {
	const styles = useStyles();
	return (
		<Dialog open={open} onOpenChange={(_event, data) => !data.open && onClose()}>
			<DialogSurface>
				<DialogBody>
					<DialogTitle>Keyboard shortcuts</DialogTitle>
					<DialogContent>
						{SECTIONS.map(({ scope, title }) => {
							const headingId = `fxs-shortcuts-${scope}`;
							return (
								<div
									key={scope}
									role="group"
									className={styles.section}
									aria-labelledby={headingId}
								>
									<Subtitle2 as="h3" id={headingId}>
										{title}
									</Subtitle2>
									<dl className={styles.list}>
										{SHORTCUTS.filter(
											(s) => s.scope === scope && !s.pending,
										).map((s) => (
											<Fragment key={s.id}>
												<dt className={styles.term}>
													<Body1>{s.label}</Body1>
												</dt>
												<dd className={styles.keys}>
													{formatKeys(s).map((key, i) => (
														<Fragment key={key}>
															{i > 0 && <span>or</span>}
															<kbd className={styles.kbd}>{key}</kbd>
														</Fragment>
													))}
												</dd>
											</Fragment>
										))}
									</dl>
								</div>
							);
						})}
					</DialogContent>
					<DialogActions>
						<Button appearance="primary" onClick={onClose}>
							Close
						</Button>
					</DialogActions>
				</DialogBody>
			</DialogSurface>
		</Dialog>
	);
}
