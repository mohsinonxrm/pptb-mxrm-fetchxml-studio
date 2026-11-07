/**
 * The target of a write operation, shown at the top of its confirmation (#69): the environment
 * (name, URL, type, with Production highlighted) and the scope (entity and record count). For
 * Production, more than 100 records, or an unknown total, it can also ask for the entity logical
 * name to be typed before the primary button enables.
 *
 * The connection comes from the host when the confirmation opens (host.getPrimaryConnection), so it
 * names the environment the requests will actually go to.
 */

import { useEffect, useState } from "react";
import {
	Badge,
	Field,
	Input,
	MessageBar,
	MessageBarBody,
	MessageBarTitle,
	Text,
	makeStyles,
	tokens,
} from "@fluentui/react-components";
import { getPrimaryConnection, type Connection } from "../../features/fetchxml/api/host";
import { metadataCache } from "../../features/fetchxml/state/cache";
import {
	describeScope,
	isProductionEnvironment,
	isTypedConfirmation,
	requiresTypedConfirmation,
	type TargetScopeKind,
} from "./targetScopeRules";

const useStyles = makeStyles({
	root: {
		display: "flex",
		flexDirection: "column",
		gap: tokens.spacingVerticalS,
	},
	box: {
		display: "grid",
		gridTemplateColumns: "auto 1fr",
		columnGap: tokens.spacingHorizontalM,
		rowGap: tokens.spacingVerticalXS,
		padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
		backgroundColor: tokens.colorNeutralBackground3,
		borderRadius: tokens.borderRadiusMedium,
	},
	label: {
		color: tokens.colorNeutralForeground3,
	},
	value: {
		display: "flex",
		flexWrap: "wrap",
		alignItems: "center",
		columnGap: tokens.spacingHorizontalS,
		minWidth: 0,
	},
	secondary: {
		color: tokens.colorNeutralForeground3,
		wordBreak: "break-all",
	},
});

type ConnectionState = { status: "loading" } | { status: "ready"; connection: Connection | null };

export interface TargetScopeProps {
	/** Logical name of the entity the operation writes to */
	entityLogicalName?: string;
	/** Display name; read from the metadata cache when not given */
	entityDisplayName?: string;
	/** What the operation applies to */
	scope: TargetScopeKind;
	/**
	 * Enables the type-to-confirm rule (requiresTypedConfirmation): called with whether the action
	 * may run. False while the connection is being read. Omit to never ask.
	 */
	onConfirmationChange?: (confirmed: boolean) => void;
	disabled?: boolean;
}

export function TargetScope({
	entityLogicalName,
	entityDisplayName,
	scope,
	onConfirmationChange,
	disabled,
}: TargetScopeProps) {
	const styles = useStyles();
	const [state, setState] = useState<ConnectionState>({ status: "loading" });
	const [typed, setTyped] = useState("");

	useEffect(() => {
		let current = true;
		getPrimaryConnection()
			.then((connection) => current && setState({ status: "ready", connection }))
			.catch((error: unknown) => {
				console.warn("TargetScope: couldn't read the connection:", error);
				if (current) setState({ status: "ready", connection: null });
			});
		return () => {
			current = false;
		};
	}, []);

	const connection = state.status === "ready" ? state.connection : null;
	const isProduction = isProductionEnvironment(connection?.environment);
	const displayName =
		entityDisplayName ||
		(entityLogicalName ? metadataCache.getEntityMetadata(entityLogicalName)?.DisplayName?.UserLocalizedLabel?.Label : undefined);
	const expected = entityLogicalName || "confirm";
	const askToType = !!onConfirmationChange && state.status === "ready" && requiresTypedConfirmation(connection?.environment, scope);
	const confirmed = state.status === "ready" && (!askToType || isTypedConfirmation(typed, expected));

	useEffect(() => {
		onConfirmationChange?.(confirmed);
	}, [confirmed, onConfirmationChange]);

	return (
		<div className={styles.root}>
			{isProduction && (
				<MessageBar intent="warning">
					<MessageBarBody>
						<MessageBarTitle>Production environment</MessageBarTitle>
						This changes live data.
					</MessageBarBody>
				</MessageBar>
			)}
			<div className={styles.box}>
				<Text className={styles.label}>Environment</Text>
				<div className={styles.value}>
					{state.status === "loading" ? (
						<Text>Reading the connection...</Text>
					) : connection ? (
						<>
							<Text weight="semibold">{connection.name}</Text>
							{connection.environment && (
								<Badge
									size="small"
									appearance={isProduction ? "filled" : "tint"}
									color={isProduction ? "warning" : "informative"}
								>
									{connection.environment}
								</Badge>
							)}
							<Text size={200} className={styles.secondary}>
								{connection.url}
							</Text>
						</>
					) : (
						<Text>Unknown (the tool has no connection)</Text>
					)}
				</div>
				<Text className={styles.label}>Entity</Text>
				<div className={styles.value}>
					<Text weight="semibold">{displayName || entityLogicalName || "Unknown"}</Text>
					{displayName && entityLogicalName && displayName !== entityLogicalName && <Text className={styles.secondary}>({entityLogicalName})</Text>}
				</div>
				<Text className={styles.label}>Scope</Text>
				<div className={styles.value}>
					<Text weight="semibold">{describeScope(scope)}</Text>
				</div>
			</div>
			{askToType && (
				<Field
					label={
						<>
							Type <strong>{expected}</strong> to confirm
						</>
					}
					hint={isProduction ? "Required on a Production environment." : "Required for more than 100 records, or when the total isn't known."}
				>
					<Input value={typed} onChange={(_, data) => setTyped(data.value)} disabled={disabled} autoComplete="off" />
				</Field>
			)}
		</div>
	);
}
