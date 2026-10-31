/**
 * Shared Relationship Picker component using Fluent UI v9 Combobox
 * Loads relationships from metadata and displays categorized by type (N:1, 1:N, N:N)
 *
 * Two-line options (#25): the related entity, then the lookup and the schema name; the search
 * matches any part.
 */

import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
	Combobox,
	Option,
	OptionGroup,
	makeStyles,
	useId,
	Spinner,
	tokens,
	type ComboboxProps,
} from "@fluentui/react-components";
import { debugLog } from "../utils/debug";
import {
	loadAllEntities,
	loadAllEntityAttributes,
	loadEntityRelationships,
} from "../../features/fetchxml/api/dataverseMetadata";
import { metadataCache } from "../../features/fetchxml/state/cache";
import { useAdvancedFindOnly } from "../../features/fetchxml/state/metadataScope";
import type { AttributeMetadata, EntityMetadata } from "../../features/fetchxml/api/pptbClient";
import { formatDataverseError } from "../../features/fetchxml/api/dataverseErrors";
import {
	RELATIONSHIP_CATEGORIES,
	buildRelationshipOptions,
	filterRelationshipOptions,
	type RelationshipOption,
} from "../../features/fetchxml/model/relationshipOptions";

export type { RelationshipOption } from "../../features/fetchxml/model/relationshipOptions";

const useStyles = makeStyles({
	option: {
		display: "flex",
		flexDirection: "column",
		minWidth: 0,
	},
	optionSecondary: {
		color: tokens.colorNeutralForeground3,
		fontSize: tokens.fontSizeBase200,
	},
});

/** Wider than the field, so most options stay on two lines (#25) */
const LISTBOX_PROPS = { style: { minWidth: 420 } };

const label = (item: { DisplayName?: { UserLocalizedLabel?: { Label?: string } | null } } | undefined) =>
	item?.DisplayName?.UserLocalizedLabel?.Label || undefined;

/**
 * Metadata the picker couldn't load (#201): the relationships themselves, or only the display names
 * (the list then shows logical names). The field around the picker says why and offers retry.
 */
export interface RelationshipLoadIssue {
	kind: "relationships" | "names";
	/** Dataverse's reason (dataverseErrors) */
	message: string;
	/** Load again; failed loads aren't cached */
	retry: () => void;
}

interface RelationshipPickerProps {
	entityLogicalName: string;
	value?: string; // SchemaName of selected relationship
	onChange: (relationship: RelationshipOption) => void;
	placeholder?: string;
	disabled?: boolean;
	/** Called when a load fails or recovers (undefined), for the surrounding field's message (#201) */
	onLoadIssue?: (issue: RelationshipLoadIssue | undefined) => void;
}

export function RelationshipPicker({
	entityLogicalName,
	value,
	onChange,
	placeholder = "Select a relationship…",
	disabled = false,
	onLoadIssue,
}: RelationshipPickerProps) {
	const styles = useStyles();
	const comboId = useId("relationship-combobox");
	const [relationships, setRelationships] = useState<RelationshipOption[]>([]);
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string | null>(null);
	// Display names that couldn't load: the list still works, with logical names (#201)
	const [namesError, setNamesError] = useState<string | null>(null);
	// Retry (#201): bumped to load again; failed loads aren't cached
	const [attempt, setAttempt] = useState(0);
	const retry = useCallback(() => setAttempt((n) => n + 1), []);
	// The latest callback, so a new function from the parent doesn't re-report the same issue
	const onLoadIssueRef = useRef(onLoadIssue);
	useEffect(() => {
		onLoadIssueRef.current = onLoadIssue;
	});
	useEffect(() => {
		onLoadIssueRef.current?.(
			error
				? { kind: "relationships", message: error, retry }
				: namesError
					? { kind: "names", message: namesError, retry }
					: undefined,
		);
	}, [error, namesError, retry]);
	const [query, setQuery] = useState<string>("");
	const [selectedId, setSelectedId] = useState<string | null>(null);
	// Reload when the "Advanced Find only" setting changes (#57)
	const advancedFindOnly = useAdvancedFindOnly();

	// Load relationships when entity changes
	useEffect(() => {
		if (!entityLogicalName) {
			debugLog("relationshipPicker", "No entity name provided, clearing relationships");
			setRelationships([]);
			return;
		}

		debugLog("relationshipPicker", "Loading relationships for entity:", entityLogicalName);
		let current = true;
		setLoading(true);
		setError(null);
		setNamesError(null);
		let namesFailure: unknown;
		const names = <T,>(load: Promise<T[]>): Promise<T[]> =>
			load.catch((failure: unknown): T[] => {
				namesFailure ??= failure;
				return [];
			});

		// Display names: every entity (cached for the entity picker), this entity's attributes (the
		// many-to-one lookups) and any related entity's attributes already cached (one-to-many lookups)
		Promise.all([
			loadEntityRelationships(entityLogicalName, advancedFindOnly),
			names<EntityMetadata>(loadAllEntities()),
			names<AttributeMetadata>(loadAllEntityAttributes(entityLogicalName)),
		])
			.then(([data, entities, ownAttributes]) => {
				if (!current) return;
				debugLog("relationshipPicker", "Received relationship data:", {
					entityLogicalName,
					oneToManyCount: data.oneToMany.length,
					manyToOneCount: data.manyToOne.length,
					manyToManyCount: data.manyToMany.length,
				});

				const entityByName = new Map(entities.map((e) => [e.LogicalName, e]));
				const attributesOf = (entity: string) =>
					entity === entityLogicalName ? ownAttributes : (metadataCache.getEntityAttributes(`${entity}__all`) ?? []);
				const options = buildRelationshipOptions(entityLogicalName, data, {
					entity: (name) => label(entityByName.get(name) ?? metadataCache.getEntityMetadata(name)),
					attribute: (entity, attribute) => label(attributesOf(entity).find((a) => a.LogicalName === attribute)),
				});

				setRelationships(options);
				setNamesError(namesFailure === undefined ? null : formatDataverseError(namesFailure, "The request failed."));
				setLoading(false);
				debugLog("relationshipPicker", "Relationships loaded successfully:", {
					entityLogicalName,
					totalRelationships: options.length,
				});
			})
			.catch((err: unknown) => {
				if (!current) return;
				console.error("[RelationshipPicker] Failed to load relationships:", {
					entityLogicalName,
					error: err,
				});
				setError(formatDataverseError(err, "The request failed."));
				setLoading(false);
			});
		return () => {
			current = false;
		};
	}, [entityLogicalName, advancedFindOnly, attempt]);

	// Sync query with selected value
	useEffect(() => {
		if (value) {
			const rel = relationships.find((r) => r.schemaName === value);
			setSelectedId(rel?.id ?? null);
			setQuery(rel ? rel.displayText : value);
		} else {
			setSelectedId(null);
			setQuery("");
		}
	}, [value, relationships]);

	// Manual filtering that preserves grouping; the selected value's own text shows every option
	const selected = relationships.find((r) => r.id === selectedId);
	const searchQuery = selected && query === selected.displayText ? "" : query;
	const filteredRelationships = useMemo(() => filterRelationshipOptions(relationships, searchQuery), [relationships, searchQuery]);

	const onOptionSelect: ComboboxProps["onOptionSelect"] = (_e, data) => {
		const id = data.optionValue;
		debugLog("relationshipPicker", "Relationship selected:", { id });

		if (id) {
			const rel = relationships.find((r) => r.id === id);
			if (rel) {
				debugLog("relationshipPicker", "Found relationship details:", rel);
				setSelectedId(rel.id);
				setQuery(rel.displayText);
				onChange(rel);
			}
		} else {
			// Handle clear
			setSelectedId(null);
			setQuery("");
		}
	};

	if (loading) {
		return <Spinner size="tiny" label="Loading relationships…" />;
	}

	if (error) {
		// The field around the picker shows why, with Retry (onLoadIssue, #201)
		// selectedOptions is required here: without it Fluent's Combobox read undefined.length and the
		// panel crashed instead of showing the error (found testing #201)
		return (
			<Combobox
				aria-labelledby={comboId}
				value=""
				selectedOptions={[]}
				placeholder="Couldn't load relationships"
				disabled={true}
			/>
		);
	}

	return (
		<Combobox
			aria-labelledby={comboId}
			listbox={LISTBOX_PROPS}
			value={query}
			selectedOptions={selectedId ? [selectedId] : []}
			onOptionSelect={onOptionSelect}
			onChange={(ev) => setQuery(ev.target.value)}
			placeholder={relationships.length === 0 ? "No relationships found" : placeholder}
			disabled={disabled || relationships.length === 0}
			clearable
		>
			{filteredRelationships.length === 0 && searchQuery.trim() ? (
				<Option disabled>{`No relationships match "${searchQuery.trim()}"`}</Option>
			) : (
				RELATIONSHIP_CATEGORIES.map((category) => {
					const options = filteredRelationships.filter((r) => r.category === category);
					return (
						options.length > 0 && (
							<OptionGroup key={category} label={category}>
								{options.map((rel) => (
									<Option key={rel.id} value={rel.id} text={rel.searchText}>
										<span className={styles.option}>
											<span>{rel.primaryText}</span>
											<span className={styles.optionSecondary}>{rel.secondaryText}</span>
										</span>
									</Option>
								))}
							</OptionGroup>
						)
					);
				})
			)}
		</Combobox>
	);
}
