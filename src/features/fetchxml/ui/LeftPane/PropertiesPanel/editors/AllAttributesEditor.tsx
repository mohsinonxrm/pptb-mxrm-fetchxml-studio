/**
 * Property panel for <all-attributes> (#90b). The element has no attributes; the panel explains
 * what it does and Microsoft's guidance against it, and offers to remove it.
 * https://learn.microsoft.com/power-apps/developer/data-platform/fetchxml/reference/all-attributes
 */

import { Body1, Button, MessageBar, MessageBarBody } from "@fluentui/react-components";
import { DOCS } from "../docLinks";
import { useEditorStyles } from "../editorStyles";
import { Info, SectionHeading } from "../fields";

interface AllAttributesEditorProps {
	isAggregateQuery: boolean;
	onRemove: () => void;
}

export function AllAttributesEditor({ isAggregateQuery, onRemove }: AllAttributesEditorProps) {
	const styles = useEditorStyles();
	return (
		<div className={styles.container}>
			<div className={styles.section}>
				<SectionHeading>All attributes</SectionHeading>
				<Body1>
					Returns every attribute that has a value, the same as listing no attributes at all.
				</Body1>
				<MessageBar intent="info" layout="multiline">
					<MessageBarBody>
						<Info
							text="Microsoft doesn't recommend this for most queries: returning every attribute is slower. Select the attributes you need instead."
							href={DOCS.antiPatterns}
						/>
					</MessageBarBody>
				</MessageBar>
				{isAggregateQuery && (
					<MessageBar intent="info" layout="multiline">
						<MessageBarBody>Ignored in an aggregate query: only the Group by and Aggregate attributes are returned.</MessageBarBody>
					</MessageBar>
				)}
				<div>
					<Button onClick={onRemove}>Remove all attributes</Button>
				</div>
			</div>
		</div>
	);
}
