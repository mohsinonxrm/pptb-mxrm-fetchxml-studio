/**
 * Value field for the hierarchy operators under, not-under, above, eq-or-under and eq-or-above
 * (#98): they take the primary key of the record to start from, so the field searches the entity's
 * records by primary name and stores the ID, with a GUID field as the fallback.
 * https://learn.microsoft.com/power-apps/developer/data-platform/query-hierarchical-data
 * ("only apply for the table relationship specifically defined as a hierarchical relationship";
 * "a default limit of 100 recursions")
 */

import { useEffect, useMemo, useRef, useState } from "react";
import { Combobox, Input, Link, Option, Spinner } from "@fluentui/react-components";
import type { EntityMetadata, RecordReference } from "../../../../api/pptbClient";
import { getHierarchicalRelationship, getRecordName, searchRecordsByName } from "../../../../api/pptbClient";
import { loadAllEntityAttributes } from "../../../../api/dataverseMetadata";
import { memoizeForSession } from "../../../../api/sessionCache";
import { isGuid, normalizeGuid, type RecordSearchTarget } from "../../../../api/recordSearch";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoField } from "../fields";

const SEARCH_DELAY_MS = 300;
const GUID_EXAMPLE = "00000000-0000-0000-0000-000000000000";
const NO_RECORDS = "__none";

/** "Hierarchy: Parent Account (parentaccountid). " for the hint; empty when unknown. */
async function hierarchyLabel(entityName: string): Promise<string> {
	const relationship = await memoizeForSession(`hierarchical-relationship:${entityName}`, () =>
		getHierarchicalRelationship(entityName),
	);
	if (!relationship) return "";
	const attribute = relationship.referencingAttribute;
	const attributes = await loadAllEntityAttributes(entityName).catch(() => []);
	const label = attributes.find((a) => a.LogicalName === attribute)?.DisplayName?.UserLocalizedLabel?.Label;
	return `Hierarchy: ${label ? `${label} (${attribute})` : attribute}. `;
}

interface HierarchyValueFieldProps {
	/** The entity whose primary key the condition is on */
	entityName: string;
	metadata: EntityMetadata | undefined;
	value: unknown;
	onChange: (value: string | undefined) => void;
}

export function HierarchyValueField({ entityName, metadata, value, onChange }: HierarchyValueFieldProps) {
	const styles = useEditorStyles();
	const target = useMemo<RecordSearchTarget | undefined>(
		() =>
			metadata?.EntitySetName && metadata.PrimaryNameAttribute && metadata.PrimaryIdAttribute
				? { entitySetName: metadata.EntitySetName, idAttribute: metadata.PrimaryIdAttribute, nameAttribute: metadata.PrimaryNameAttribute }
				: undefined,
		[metadata],
	);
	const stored = typeof value === "string" ? value : "";
	// The value this field last wrote: only a different value (Parse to Tree, undo) is read back
	const written = useRef<string | null>(null);
	const [byId, setById] = useState(false);
	const [typed, setTyped] = useState(stored);
	const [query, setQuery] = useState("");
	const [selectedName, setSelectedName] = useState<string | undefined>(undefined);
	const [results, setResults] = useState<RecordReference[] | undefined>(undefined);
	const [searching, setSearching] = useState(false);
	const [notFound, setNotFound] = useState(false);
	const [hierarchy, setHierarchy] = useState("");
	const searchRun = useRef(0);

	useEffect(() => {
		let cancelled = false;
		hierarchyLabel(entityName).then(
			(text) => !cancelled && setHierarchy(text),
			() => !cancelled && setHierarchy(""),
		);
		return () => {
			cancelled = true;
		};
	}, [entityName]);

	// The stored ID's name (Parse to Tree, another condition); an ID that can't be read opens as an ID.
	// A run counter, not an effect cleanup, drops stale answers: the refs make a repeated run (Strict
	// Mode) return early, and a cleanup would cancel the only request.
	const resolvedWith = useRef<RecordSearchTarget | undefined>(undefined);
	const resolveRun = useRef(0);
	useEffect(() => {
		if (stored === written.current && target === resolvedWith.current) return;
		written.current = stored;
		resolvedWith.current = target;
		const run = ++resolveRun.current;
		setTyped(stored);
		setNotFound(false);
		if (!target || !stored) {
			setSelectedName(undefined);
			setQuery("");
			return;
		}
		if (!isGuid(stored)) {
			setById(true);
			return;
		}
		getRecordName(target, stored).then((name) => {
			if (run !== resolveRun.current) return;
			if (name === undefined) {
				setNotFound(true);
				setById(true);
				return;
			}
			setSelectedName(name);
			setQuery(name || normalizeGuid(stored));
		});
	}, [stored, target]);

	const write = (next: string | undefined) => {
		written.current = next ?? "";
		// A pending name lookup for the previous value no longer applies
		resolveRun.current++;
		onChange(next);
	};

	const search = (text: string) => {
		if (!target) return;
		const run = ++searchRun.current;
		setSearching(true);
		searchRecordsByName(target, text).then(
			(records) => {
				if (run !== searchRun.current) return;
				setResults(records);
				setSearching(false);
			},
			(error: unknown) => {
				if (run !== searchRun.current) return;
				console.error("Record search failed:", error);
				setResults([]);
				setSearching(false);
			},
		);
	};

	// Search as the user types, after a pause
	const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
	useEffect(() => () => clearTimeout(timer.current), []);
	const onQueryChange = (text: string) => {
		setQuery(text);
		clearTimeout(timer.current);
		timer.current = setTimeout(() => search(text), SEARCH_DELAY_MS);
	};

	const entityLabel = metadata?.DisplayName?.UserLocalizedLabel?.Label ?? entityName;
	const hint = `${hierarchy}Dataverse follows up to 100 levels.`;
	const info = (
		<Info
			text={`The ${entityLabel} record to start from. Its ID (${metadata?.PrimaryIdAttribute ?? "primary key"}) is what the condition stores. The operators apply only to a hierarchical relationship.`}
			href={DOCS.hierarchy}
		/>
	);
	const toggle = target && (
		<Link
			as="button"
			inline
			onClick={() => {
				if (!byId) setTyped(stored);
				setById((b) => !b);
			}}
		>
			{byId ? "Search for a record instead" : "Enter an ID instead"}
		</Link>
	);

	if (byId || !target) {
		const invalid = typed.trim() !== "" && !isGuid(typed);
		return (
			<>
				<InfoField
					label="Record ID"
					info={info}
					hint={notFound ? `No ${entityLabel} record with this ID was found, or you can't read it. ${hint}` : hint}
					validation={invalid ? { state: "error", message: `Enter an ID like ${GUID_EXAMPLE}.` } : undefined}
				>
					<Input
						className={styles.mono}
						value={typed}
						placeholder={GUID_EXAMPLE}
						onChange={(_, data) => {
							setTyped(data.value);
							const text = data.value.trim();
							write(text === "" ? undefined : isGuid(text) ? normalizeGuid(text) : text);
						}}
					/>
				</InfoField>
				{toggle}
			</>
		);
	}

	return (
		<>
			<InfoField label="Record" info={info} hint={hint}>
				<Combobox
					freeform
					placeholder="Search by name (begins with)"
					value={query}
					selectedOptions={stored ? [normalizeGuid(stored).toLowerCase()] : []}
					onChange={(ev) => onQueryChange(ev.target.value)}
					onOpenChange={(_, data) => {
						if (data.open && results === undefined) search(query === selectedName ? "" : query);
					}}
					onOptionSelect={(_, data) => {
						if (!data.optionValue || data.optionValue === NO_RECORDS) return;
						const record = results?.find((r) => r.id.toLowerCase() === data.optionValue);
						setSelectedName(record?.name);
						setQuery(record?.name || data.optionValue);
						write(data.optionValue);
					}}
				>
					{searching && results === undefined ? (
						<Option disabled value={NO_RECORDS} text="">
							<Spinner size="tiny" label="Searching…" />
						</Option>
					) : results && results.length > 0 ? (
						results.map((record) => (
							<Option key={record.id} value={record.id.toLowerCase()} text={record.name || record.id}>
								<span className={styles.optionText}>
									{record.name || "(No name)"}
									<span className={styles.optionSecondary}>{record.id}</span>
								</span>
							</Option>
						))
					) : (
						<Option disabled value={NO_RECORDS} text="">
							{query ? `No ${entityLabel} records start with "${query}"` : `No ${entityLabel} records found`}
						</Option>
					)}
				</Combobox>
			</InfoField>
			{toggle}
		</>
	);
}
