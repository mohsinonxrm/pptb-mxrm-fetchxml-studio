/**
 * Shared Relationship Picker component using Fluent UI v9 Combobox
 * Loads relationships from metadata and displays categorized by type (1:N, N:1, N:N)
 */

import { useState, useEffect } from "react";
import {
	Combobox,
	Option,
	OptionGroup,
	useId,
	Spinner,
	type ComboboxProps,
} from "@fluentui/react-components";
import { debugLog } from "../utils/debug";
import { loadEntityRelationships } from "../../features/fetchxml/api/dataverseMetadata";
import type { RelationshipMetadata } from "../../features/fetchxml/api/pptbClient";

interface RelationshipOption {
	schemaName: string;
	displayText: string;
	referencedEntity: string;
	referencingEntity: string;
	referencedAttribute: string;
	referencingAttribute: string;
	type: "1N" | "N1" | "NN";
	category: string;
}

interface RelationshipPickerProps {
	entityLogicalName: string;
	value?: string; // SchemaName of selected relationship
	onChange: (relationship: RelationshipOption) => void;
	placeholder?: string;
	disabled?: boolean;
}

export function RelationshipPicker({
	entityLogicalName,
	value,
	onChange,
	placeholder = "Select a relationship",
	disabled = false,
}: RelationshipPickerProps) {
	const comboId = useId("relationship-combobox");
	const [relationships, setRelationships] = useState<RelationshipOption[]>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string | null>(null);
	const [query, setQuery] = useState<string>("");

	// Load relationships when entity changes
	useEffect(() => {
		if (!entityLogicalName) {
			debugLog("relationshipPicker", "No entity name provided, clearing relationships");
			setRelationships([]);
			return;
		}

		debugLog("relationshipPicker", "Loading relationships for entity:", entityLogicalName);
		setLoading(true);
		setError(null);

		loadEntityRelationships(entityLogicalName)
			.then((data) => {
				debugLog("relationshipPicker", "Received relationship data:", {
					entityLogicalName,
					oneToManyCount: data.oneToMany.length,
					manyToOneCount: data.manyToOne.length,
					manyToManyCount: data.manyToMany.length,
				});

				const options: RelationshipOption[] = [];

				// Process 1:N relationships (current entity is Referenced/parent)
				data.oneToMany.forEach((rel: RelationshipMetadata) => {
					// Show: childEntity → parentEntity via lookupField (SchemaName)
					const lookupField = rel.ReferencingAttribute || "";
					const displayText = lookupField
						? `${rel.ReferencingEntity} → ${rel.ReferencedEntity} via ${lookupField} (${rel.SchemaName})`
						: `${rel.ReferencingEntity} → ${rel.ReferencedEntity} (${rel.SchemaName})`;

					options.push({
						schemaName: rel.SchemaName,
						displayText,
						referencedEntity: rel.ReferencedEntity,
						referencingEntity: rel.ReferencingEntity,
						referencedAttribute: rel.ReferencedAttribute,
						referencingAttribute: rel.ReferencingAttribute,
						type: "1N",
						category: "One-to-Many (1:N)",
					});
				});

				// Process N:1 relationships (current entity is Referencing/child)
				data.manyToOne.forEach((rel: RelationshipMetadata) => {
					// Show: childEntity → parentEntity via lookupField (SchemaName)
					const lookupField = rel.ReferencingAttribute || "";
					const displayText = lookupField
						? `${rel.ReferencingEntity} → ${rel.ReferencedEntity} via ${lookupField} (${rel.SchemaName})`
						: `${rel.ReferencingEntity} → ${rel.ReferencedEntity} (${rel.SchemaName})`;

					options.push({
						schemaName: rel.SchemaName,
						displayText,
						referencedEntity: rel.ReferencedEntity,
						referencingEntity: rel.ReferencingEntity,
						referencedAttribute: rel.ReferencedAttribute,
						referencingAttribute: rel.ReferencingAttribute,
						type: "N1",
						category: "Many-to-One (N:1)",
					});
				});

				// Process N:N relationships
				data.manyToMany.forEach((rel: RelationshipMetadata) => {
					// N:N relationships have Entity1/Entity2, not Referencing/Referenced
					const entity1 = rel.Entity1LogicalName || "";
					const entity2 = rel.Entity2LogicalName || "";
					const attr1 = rel.Entity1IntersectAttribute || "";
					const attr2 = rel.Entity2IntersectAttribute || "";

					options.push({
						schemaName: rel.SchemaName,
						displayText: `${entity1} ↔ ${entity2} (${rel.SchemaName})`,
						referencedEntity: entity2, // Store entity2 as "referenced" for compatibility
						referencingEntity: entity1, // Store entity1 as "referencing" for compatibility
						referencedAttribute: attr2,
						referencingAttribute: attr1,
						type: "NN",
						category: "Many-to-Many (N:N)",
					});
				});

				setRelationships(options);
				setLoading(false);
				debugLog("relationshipPicker", "Relationships loaded successfully:", {
					entityLogicalName,
					totalRelationships: options.length,
				});
			})
			.catch((err: unknown) => {
				console.error("[RelationshipPicker] Failed to load relationships:", {
					entityLogicalName,
					error: err,
				});
				setError("Failed to load relationships");
				setLoading(false);
			});
	}, [entityLogicalName]);

	// Sync query with selected value
	useEffect(() => {
		if (value) {
			const rel = relationships.find((r) => r.schemaName === value);
			if (rel) {
				setQuery(rel.displayText);
			} else {
				setQuery(value);
			}
		} else {
			setQuery("");
		}
	}, [value, relationships]);

	// Manual filtering that preserves grouping
	const filteredRelationships = query.trim()
		? relationships.filter(
				(rel) =>
					rel.displayText.toLowerCase().includes(query.toLowerCase()) ||
					rel.schemaName.toLowerCase().includes(query.toLowerCase()) ||
					rel.referencingAttribute.toLowerCase().includes(query.toLowerCase())
		  )
		: relationships;

	// Group by category
	const categories = {
		"One-to-Many (1:N)": filteredRelationships.filter((r) => r.category === "One-to-Many (1:N)"),
		"Many-to-One (N:1)": filteredRelationships.filter((r) => r.category === "Many-to-One (N:1)"),
		"Many-to-Many (N:N)": filteredRelationships.filter((r) => r.category === "Many-to-Many (N:N)"),
	};

	const onOptionSelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		const schemaName = data.optionValue;
		debugLog("relationshipPicker", "Relationship selected:", {
			schemaName,
			optionText: data.optionText,
		});

		if (schemaName) {
			const rel = relationships.find((r) => r.schemaName === schemaName);
			if (rel) {
				debugLog("relationshipPicker", "Found relationship details:", rel);
				setQuery(data.optionText ?? "");
				onChange(rel);
			}
		} else {
			// Handle clear
			setQuery("");
		}
	};

	if (loading) {
		return <Spinner size="tiny" label="Loading relationships..." />;
	}

	if (error) {
		return (
			<Combobox
				aria-labelledby={comboId}
				value=""
				placeholder="Error loading relationships"
				disabled={true}
			/>
		);
	}

	return (
		<Combobox
			aria-labelledby={comboId}
			value={query}
			onOptionSelect={onOptionSelect}
			onChange={(ev) => setQuery(ev.target.value)}
			placeholder={
				loading
					? "Loading relationships..."
					: relationships.length === 0
					? "No relationships found"
					: placeholder
			}
			disabled={disabled || loading || relationships.length === 0}
			clearable
		>
			{filteredRelationships.length === 0 && query.trim() ? (
				<Option disabled>No relationships match your search</Option>
			) : (
				<>
					{categories["One-to-Many (1:N)"].length > 0 && (
						<OptionGroup label="One-to-Many (1:N)">
							{categories["One-to-Many (1:N)"].map((rel) => (
								<Option key={rel.schemaName} value={rel.schemaName}>
									{rel.displayText}
								</Option>
							))}
						</OptionGroup>
					)}
					{categories["Many-to-One (N:1)"].length > 0 && (
						<OptionGroup label="Many-to-One (N:1)">
							{categories["Many-to-One (N:1)"].map((rel) => (
								<Option key={rel.schemaName} value={rel.schemaName}>
									{rel.displayText}
								</Option>
							))}
						</OptionGroup>
					)}
					{categories["Many-to-Many (N:N)"].length > 0 && (
						<OptionGroup label="Many-to-Many (N:N)">
							{categories["Many-to-Many (N:N)"].map((rel) => (
								<Option key={rel.schemaName} value={rel.schemaName}>
									{rel.displayText}
								</Option>
							))}
						</OptionGroup>
					)}
				</>
			)}
		</Combobox>
	);
}
