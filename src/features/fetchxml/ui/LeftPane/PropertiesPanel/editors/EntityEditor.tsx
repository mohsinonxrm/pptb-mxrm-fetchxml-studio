/**
 * Property editor for the root <entity>: its logical name (changed with the Entity picker) and the
 * report prefiltering attributes (#90b).
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/entity
 */

import {
	Accordion,
	AccordionHeader,
	AccordionItem,
	AccordionPanel,
	Badge,
	Input,
	Text,
} from "@fluentui/react-components";
import type { EntityNode } from "../../../../model/nodes";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, InfoCheckbox, InfoField } from "../fields";
import { useEntityInfo } from "../useEditorMetadata";
import { MetadataLoadWarning } from "../MetadataLoadWarning";

interface EntityEditorProps {
	node: EntityNode;
	onUpdate: (updates: Record<string, unknown>) => void;
}

export function EntityEditor({ node, onUpdate }: EntityEditorProps) {
	const styles = useEditorStyles();
	const { metadata } = useEntityInfo(node.name);
	const displayName = metadata?.DisplayName?.UserLocalizedLabel?.Label;
	const parameterName = node.prefilterParameterName ?? "";
	const prefilteringSet = !!node.enablePrefiltering || parameterName !== "";

	return (
		<div className={styles.container}>
			{/* Metadata that couldn't load, with Retry (#201) */}
			<MetadataLoadWarning entityName={node.name} />
			<div className={styles.section}>
				<InfoField
					label="Entity"
					hint="Change it with the Entity picker above the tree."
					info={<Info text="The logical name of the entity the query is based on." href={DOCS.entity} />}
				>
					<Text className={styles.readOnly}>
						{node.name}
						{displayName && <span className={styles.secondary}> ({displayName})</span>}
					</Text>
				</InfoField>
			</div>

			<Accordion collapsible defaultOpenItems={prefilteringSet ? ["prefilter"] : []}>
				<AccordionItem value="prefilter">
					<AccordionHeader>
						<span className={styles.accordionHeader}>
							Report prefiltering
							{prefilteringSet && (
								<Badge appearance="tint" size="small" color="informative">
									set
								</Badge>
							)}
						</span>
					</AccordionHeader>
					<AccordionPanel>
						<div className={styles.accordionPanel}>
							<InfoCheckbox
								label="Enable prefiltering"
								checked={!!node.enablePrefiltering}
								onChange={(checked) => onUpdate({ enablePrefiltering: checked ? true : undefined })}
								info={
									<Info
										text="For SQL Server Reporting Services (Fetch-based) reports: lets people filter the report's records with Advanced Find before it runs. Dataverse ignores these attributes when the query runs here."
										href={DOCS.reportPrefiltering}
									/>
								}
							/>
							{(node.enablePrefiltering || parameterName) && (
								<InfoField
									label="Parameter name"
									validation={
										parameterName && !parameterName.startsWith("CRM_")
											? { state: "warning", message: "Start the name with CRM_ so the report hides the parameter." }
											: undefined
									}
									info={<Info text="The report parameter that receives the prefilter." href={DOCS.reportPrefiltering} />}
								>
									<Input
										className={styles.mono}
										value={parameterName}
										placeholder="CRM_FilteredAccount"
										onChange={(_, data) => onUpdate({ prefilterParameterName: data.value || undefined })}
									/>
								</InfoField>
							)}
						</div>
					</AccordionPanel>
				</AccordionItem>
			</Accordion>
		</div>
	);
}
