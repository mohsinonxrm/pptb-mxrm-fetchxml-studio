/**
 * "Couldn't load columns for account" with Dataverse's reason and Retry (#201), at the top of a
 * property editor whose entity metadata failed to load. The editor keeps working with free text.
 */
import { Button, MessageBar, MessageBarActions, MessageBarBody } from "@fluentui/react-components";
import { BarTitle } from "../../../../../shared/components/BarTitle";
import { retryMetadataLoads, useMetadataLoadError } from "./metadataLoadErrors";

export function MetadataLoadWarning({ entityName }: { entityName: string | undefined }) {
	const failure = useMetadataLoadError(entityName);
	if (!failure || !entityName) return null;
	const what = failure.kind === "columns" ? "attributes" : "entity details";
	return (
		<MessageBar intent="warning" layout="multiline">
			<MessageBarBody>
				<BarTitle>{`Couldn't load ${what} for ${entityName}`}</BarTitle>
				{failure.message}
			</MessageBarBody>
			<MessageBarActions>
				<Button size="small" onClick={retryMetadataLoads}>
					Retry
				</Button>
			</MessageBarActions>
		</MessageBar>
	);
}
