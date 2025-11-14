/**
 * Entity selector dropdown for choosing the root query entity
 */

import { useState, useEffect } from "react";
import {
	Combobox,
	makeStyles,
	Button,
	Spinner,
	useId,
	useComboboxFilter,
	tokens,
	type ComboboxProps,
} from "@fluentui/react-components";
import { Add20Regular } from "@fluentui/react-icons";
import { useLazyMetadata } from "../../../../shared/hooks/useLazyMetadata";
import type { EntityMetadata } from "../../api/pptbClient";

const useStyles = makeStyles({
	container: {
		display: "flex",
		alignItems: "flex-end",
		gap: "12px",
		padding: "8px 12px",
		borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
	},
	entityField: {
		display: "flex",
		flexDirection: "column",
		gap: "4px",
		flex: 1,
	},
	label: {
		fontSize: "14px",
		fontWeight: 600,
		color: tokens.colorNeutralForeground2,
	},
});

interface EntityOption {
	logicalName: string;
	displayName: string;
}

interface EntitySelectorProps {
	selectedEntity: string | null;
	onEntityChange: (entityLogicalName: string) => void;
	onNewQuery: () => void;
}

export function EntitySelector({
	selectedEntity,
	onEntityChange,
	onNewQuery,
}: EntitySelectorProps) {
	const styles = useStyles();
	const comboId = useId("entity-combobox");
	const [entities, setEntities] = useState<EntityOption[]>([]);
	const [isLoading, setIsLoading] = useState(true);
	const [query, setQuery] = useState<string>("");
	const { loadEntities } = useLazyMetadata();

	// Update query when selectedEntity changes
	useEffect(() => {
		if (selectedEntity) {
			const entity = entities.find((e) => e.logicalName === selectedEntity);
			if (entity) {
				setQuery(entity.displayName);
			}
		}
	}, [selectedEntity, entities]);

	// Load entities from PPTB
	useEffect(() => {
		loadEntities(true)
			.then((entityMetadata: EntityMetadata[]) => {
				const options = entityMetadata.map((entity) => ({
					logicalName: entity.LogicalName,
					displayName: entity.DisplayName?.UserLocalizedLabel?.Label || entity.LogicalName,
				}));
				// Sort by display name
				options.sort((a, b) => a.displayName.localeCompare(b.displayName));

				setEntities(options);
			})
			.catch((err) => {
				console.error("EntitySelector: Failed to load entities:", err);
				// Fall back to sample entities on error
				const fallbackEntities = [
					{ logicalName: "account", displayName: "Account" },
					{ logicalName: "contact", displayName: "Contact" },
					{ logicalName: "opportunity", displayName: "Opportunity" },
					{ logicalName: "lead", displayName: "Lead" },
					{ logicalName: "systemuser", displayName: "User" },
				];
				setEntities(fallbackEntities);
			})
			.finally(() => {
				setIsLoading(false);
			});
	}, [loadEntities]);

	// Filter options based on query
	const options = entities.map((entity) => ({
		children: entity.displayName,
		value: entity.logicalName,
	}));

	const filteredChildren = useComboboxFilter(query, options, {
		noOptionsMessage: "No entities match your search.",
	});

	const onOptionSelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		if (data.optionValue) {
			onEntityChange(data.optionValue);
			setQuery(data.optionText ?? "");
		} else {
			// Clear button clicked - notify parent to clear the entity (which resets the entire query)
			setQuery("");
			onEntityChange("");
		}
	};

	return (
		<div className={styles.container}>
			{isLoading ? (
				<Spinner size="tiny" label="Loading entities..." />
			) : (
				<div className={styles.entityField}>
					<label id={comboId} className={styles.label}>
						Entity
					</label>
					<Combobox
						aria-labelledby={comboId}
						placeholder="Select an entity..."
						value={query}
						onOptionSelect={onOptionSelect}
						onChange={(ev) => setQuery(ev.target.value)}
						clearable
					>
						{filteredChildren}
					</Combobox>
				</div>
			)}
			<Button
				appearance="primary"
				icon={<Add20Regular />}
				onClick={onNewQuery}
				title="Create a new query"
			>
				New
			</Button>
		</div>
	);
}
