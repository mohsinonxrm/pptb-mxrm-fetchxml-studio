/**
 * Property editor for <filter>: type, hint (union), isquickfindfields and the quick find record
 * limit overrides (#90b). https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/filter
 */

import { Dropdown, Option, Radio, RadioGroup } from "@fluentui/react-components";
import type { FilterNode } from "../../../../model/nodes";
import type { RuleFinding } from "../../../../model/fetchxmlRules";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoCheckbox, InfoField, SectionHeading } from "../fields";
import { findingFor } from "../ruleFindings";

interface FilterEditorProps {
	node: FilterNode;
	/** Another filter in the query already uses the union hint (one per query) */
	unionElsewhere: boolean;
	/** Query-rule findings for this node (fetchxmlRules.ts) */
	findings: readonly RuleFinding[];
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function FilterEditor({ node, unionElsewhere, findings, onUpdate }: FilterEditorProps) {
	const styles = useEditorStyles();
	const isOr = node.conjunction === "or";
	const anyQuickFind =
		!!node.isQuickFindFields || !!node.overrideQuickFindRecordLimitEnabled || !!node.overrideQuickFindRecordLimitDisabled;
	const recordLimit =
		node.overrideQuickFindRecordLimitEnabled && node.overrideQuickFindRecordLimitDisabled
			? ""
			: node.overrideQuickFindRecordLimitEnabled
				? "apply"
				: node.overrideQuickFindRecordLimitDisabled
					? "bypass"
					: "default";
	const unionBlocked = !node.hint && unionElsewhere;

	return (
		<div className={styles.container}>
			<div className={styles.section}>
				<SectionHeading>Filter</SectionHeading>
				<InfoField
					label="Match"
					info={<Info text="Whether all conditions (and) or any condition (or) must be true." href={DOCS.filter} />}
				>
					<Dropdown
						value={isOr ? "Any condition (or)" : "All conditions (and)"}
						selectedOptions={[node.conjunction]}
						onOptionSelect={(_, data) => onUpdate({ conjunction: data.optionValue })}
					>
						<Option value="and">All conditions (and)</Option>
						<Option value="or">Any condition (or)</Option>
					</Dropdown>
				</InfoField>

				{(isOr || node.hint) && (
					<InfoCheckbox
						label="Union hint"
						checked={!!node.hint}
						disabled={unionBlocked}
						onChange={(checked) => onUpdate({ hint: checked ? "union" : undefined })}
						hint={unionBlocked ? "Only one filter per query can use the union hint, and another filter has it." : undefined}
						validation={findingFor(findings, "hint")}
						info={
							<Info
								text="Can improve performance for an 'any condition' filter whose conditions are on attributes of different entities. One per query."
								href={DOCS.unionHint}
							/>
						}
					/>
				)}
			</div>

			{(isOr || anyQuickFind) && (
				<div className={styles.section}>
					<SectionHeading>Quick find</SectionHeading>
					<InfoCheckbox
						label="Quick find filter"
						checked={!!node.isQuickFindFields}
						onChange={(checked) => onUpdate({ isQuickFindFields: checked ? true : undefined })}
						info={
							<Info
								text="Runs this filter as a quick find search, the pattern behind the search box in views. Dataverse applies the quick find record limit to it."
								href={DOCS.quickFind}
							/>
						}
					/>
					{anyQuickFind && (
						<InfoField
							label="Record limit"
							validation={findingFor(findings, "overridequickfindrecordlimitenabled")}
							info={
								<Info
									text="When the 'Quick Find record limits' environment setting is on (the default, and recommended), a quick find matching more than 10,000 records fails with QuickFindQueryRecordLimitExceeded. The count is taken before security filters. The overrides are for temporary testing."
									href={DOCS.quickFindLimits}
								/>
							}
						>
							<RadioGroup
								value={recordLimit}
								onChange={(_, data) =>
									onUpdate({
										overrideQuickFindRecordLimitEnabled: data.value === "apply" ? true : undefined,
										overrideQuickFindRecordLimitDisabled: data.value === "bypass" ? true : undefined,
									})
								}
							>
								<Radio value="default" label="Use the environment setting (default)" />
								<Radio value="apply" label="Always apply the 10,000-record limit (testing)" />
								<Radio value="bypass" label="Bypass the limit (testing)" />
							</RadioGroup>
						</InfoField>
					)}
				</div>
			)}
		</div>
	);
}
