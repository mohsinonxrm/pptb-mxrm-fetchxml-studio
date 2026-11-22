/**
 * The inspector drawer (#234): Properties, when the window or the builder pane is too narrow for
 * them below the tree. A non-modal Fluent OverlayDrawer at the end of the output area, under the
 * tabs (they stay visible and usable), so the results can still be read and used beside it.
 * Escape or Close closes it; the caller returns focus to "Edit properties".
 */
import { useState, type ReactNode } from "react";
import {
	Button,
	DrawerBody,
	DrawerHeader,
	DrawerHeaderTitle,
	OverlayDrawer,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { Dismiss24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
	// Where the drawer mounts: the output area, so it opens over the tab's content, not the window
	host: {
		position: "absolute",
		inset: 0,
		pointerEvents: "none",
		overflow: "hidden",
	},
	drawer: {
		pointerEvents: "auto",
		borderLeft: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
	},
});

export interface InspectorDrawerProps {
	open: boolean;
	/** "Kind · name" of the node being edited */
	title: string;
	onClose: () => void;
	children: ReactNode;
}

export function InspectorDrawer({ open, title, onClose, children }: InspectorDrawerProps) {
	const styles = useStyles();
	const [host, setHost] = useState<HTMLDivElement | null>(null);
	return (
		<>
			<div ref={setHost} className={styles.host} />
			{host && (
				<OverlayDrawer
					className={styles.drawer}
					open={open}
					mountNode={host}
					position="end"
					size="small"
					modalType="non-modal"
					aria-label={`Properties: ${title}`}
					onOpenChange={(_, data) => {
						if (!data.open) onClose();
					}}
				>
					<DrawerHeader>
						<DrawerHeaderTitle
							action={
								<Button
									appearance="subtle"
									aria-label="Close properties"
									icon={<Dismiss24Regular />}
									onClick={onClose}
								/>
							}
						>
							{title}
						</DrawerHeaderTitle>
					</DrawerHeader>
					<DrawerBody>{children}</DrawerBody>
				</OverlayDrawer>
			)}
		</>
	);
}
