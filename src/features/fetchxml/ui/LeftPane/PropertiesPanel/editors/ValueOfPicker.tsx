/**
 * Picks the attribute a condition compares to (`valueof`): first the entity, then a compatible
 * attribute. An unqualified name is an attribute of the condition's own entity; other entities are
 * offered only where Dataverse can reach them from this condition (valueofScopes, verified live):
 * the root entity by name, and joined link-entities by alias.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/filter-rows#filter-on-column-values-in-the-same-row
 */

import { useEffect, useMemo, useState } from "react";
import { Dropdown, Option } from "@fluentui/react-components";
import type { FetchNode } from "../../../../model/nodes";
import type { ValueOfScope } from "../../../../model/fetchxmlRules";
import { collectLinkEntityReferences } from "../../../../model/treeUtils";
import { AttributePicker } from "../../../../../../shared/components/AttributePicker";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoField, type Validation } from "../fields";

const SAME = "__same";

interface ValueOfPickerProps {
	fetchQuery: FetchNode | null;
	rootEntityName: string;
	/** The entity the condition is on (its link-entity, or the one named by entityname) */
	effectiveEntityName: string;
	/** What this condition can reach (valueofScopes) */
	scope: ValueOfScope | undefined;
	/** Alias that means the same record as "same record" (not listed twice) */
	sameAlias?: string;
	valueof: string;
	/** AttributeTypes Dataverse can compare with the condition's attribute */
	compatibleTypes: string[];
	validation?: Validation;
	onChange: (valueof: string | undefined) => void;
}

/** "createdon" → same entity; "acct.name" → alias "acct" */
function parseValueOf(valueof: string): { qualifier: string | undefined; attribute: string } {
	const dot = valueof.indexOf(".");
	return dot === -1
		? { qualifier: undefined, attribute: valueof }
		: { qualifier: valueof.substring(0, dot), attribute: valueof.substring(dot + 1) };
}

export function ValueOfPicker({
	fetchQuery,
	rootEntityName,
	effectiveEntityName,
	scope,
	sameAlias,
	valueof,
	compatibleTypes,
	validation,
	onChange,
}: ValueOfPickerProps) {
	const styles = useEditorStyles();
	const parsed = useMemo(() => parseValueOf(valueof), [valueof]);

	const options = useMemo(() => {
		const list: { value: string; label: string; entityName: string; secondary?: string }[] = [
			{ value: SAME, label: `${effectiveEntityName} (same record)`, entityName: effectiveEntityName },
		];
		if (rootEntityName && rootEntityName !== effectiveEntityName) {
			list.push({ value: rootEntityName, label: `${rootEntityName} (root entity)`, entityName: rootEntityName });
		}
		const entityOf = new Map(collectLinkEntityReferences(fetchQuery).map((r) => [r.alias, r.entityName]));
		for (const alias of scope?.aliases ?? []) {
			if (alias === sameAlias) continue;
			list.push({ value: alias, label: `${alias} (${entityOf.get(alias) ?? "?"})`, entityName: entityOf.get(alias) ?? "" });
		}
		// Keep a parsed qualifier visible even when it can't be resolved (the rule finding explains why)
		if (parsed.qualifier && !list.some((o) => o.value === parsed.qualifier)) {
			const reachable = parsed.qualifier === sameAlias;
			list.push({
				value: parsed.qualifier,
				label: parsed.qualifier,
				entityName: reachable ? effectiveEntityName : "",
				secondary: reachable ? "This record" : scope?.laterAliases.includes(parsed.qualifier) ? "Joined after this link-entity" : "Not a link-entity alias",
			});
		}
		return list;
	}, [fetchQuery, rootEntityName, effectiveEntityName, scope, sameAlias, parsed.qualifier]);

	const [selected, setSelected] = useState<string>(parsed.qualifier ?? SAME);
	// Follow external changes (Parse to Tree, undo)
	useEffect(() => setSelected(parsed.qualifier ?? SAME), [parsed.qualifier]);

	const current = options.find((o) => o.value === selected) ?? options[0];

	return (
		<>
			<InfoField
				label="Entity"
				info={
					<Info
						text="Compare to an attribute of the same record, the root entity, or a link-entity this condition can reach: link-entities need an alias, and a link-entity's filter can only reach link-entities joined before it."
						href={DOCS.filterOnColumnValues}
					/>
				}
			>
				<Dropdown
					value={current.label}
					selectedOptions={[current.value]}
					onOptionSelect={(_, data) => {
						setSelected(data.optionValue ?? SAME);
						// Keep valueof defined (empty) so the editor stays in "Another attribute" mode
						onChange("");
					}}
				>
					{options.map((o) => (
						<Option key={o.value} value={o.value} text={o.label}>
							<span className={styles.optionText}>
								{o.label}
								{o.secondary && <span className={styles.optionSecondary}>{o.secondary}</span>}
							</span>
						</Option>
					))}
				</Dropdown>
			</InfoField>

			<InfoField
				label="Attribute"
				required
				validation={validation ?? (!parsed.attribute ? { state: "error", message: "Choose the attribute to compare to." } : undefined)}
				info={
					<Info
						text={
							compatibleTypes.length > 0
								? `Only attributes Dataverse can compare with this one are listed (${compatibleTypes.join(", ")}).`
								: "The attribute to compare to. Both attributes must have compatible types."
						}
						href={DOCS.filterOnColumnValues}
					/>
				}
			>
				<AttributePicker
					entityLogicalName={current.entityName}
					value={parsed.attribute}
					onChange={(logicalName) => {
						const qualifier = current.value === SAME ? undefined : current.value;
						onChange(logicalName ? (qualifier ? `${qualifier}.${logicalName}` : logicalName) : "");
					}}
					placeholder={current.entityName ? "Select an attribute" : "Choose an entity first"}
					disabled={!current.entityName || compatibleTypes.length === 0}
					filterByTypes={compatibleTypes}
				/>
			</InfoField>
		</>
	);
}
