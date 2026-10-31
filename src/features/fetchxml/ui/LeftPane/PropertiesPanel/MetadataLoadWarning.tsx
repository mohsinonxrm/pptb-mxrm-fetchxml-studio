/**
 * "Couldn't load columns for account" with Dataverse's reason and Retry (#201), at the top of a
 * property editor whose entity metadata failed to load. The editor keeps working with free text.
 */
import { Button, MessageBar, MessageBarActions, MessageBarBody, MessageBarTitle } from "@fluentui/react-components";
import { retryMetadataLoads, useMetadataLoadError } from "./metadataLoadErrors";

export function MetadataLoadWarning({ entityName }: { entityName: string | undefined }) {
	const failure = useMetadataLoadError(entityName);
	if (!failure || !entityName) return null;
	const what = failure.kind === "columns" ? "columns" : "entity details";
	return (
		<MessageBar intent="warning" layout="multiline">
			<MessageBarBody>
				<MessageBarTitle>{`Couldn't load ${what} for ${entityName}`}</MessageBarTitle>
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
