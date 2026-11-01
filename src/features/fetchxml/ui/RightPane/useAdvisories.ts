/**
 * Advisories tab state (#97): the metadata the metadata rules need, from the metadata cache, and
 * the on-demand "Check with Dataverse" (ValidateFetchXmlExpression).
 */
import { useCallback, useEffect, useRef, useState } from "react";
import { isDataverseAvailable, validateFetchXmlExpression } from "../../api/pptbClient";
import { loadAllEntityAttributes, loadEntityMetadata, loadStringMaxLengths } from "../../api/dataverseMetadata";
import { formatDataverseError } from "../../api/dataverseErrors";
import { toAttributeTypeInfo } from "../../model/attributeCapabilities";
import { advisorEntityNames, type DataverseCheck } from "../../model/advisories";
import type { AdvisorAttributeInfo, AdvisorEntityInfo, AdvisorMetadata } from "../../model/queryAdvisor";
import type { FetchNode } from "../../model/nodes";

/**
 * Metadata for the metadata rules: attribute types and SourceType for every entity in the query,
 * String MaxLength, and the root's primary key. Undefined parts mean "not loaded" (still loading,
 * no connection, or the request failed), and the rules that need them are skipped.
 *
 * Loaded only once `enabled` (#170): from the Advisories tab's first open, not as soon as an entity
 * is chosen (in the spirit of #71). The tab's count then grows when it's opened.
 */
export function useAdvisorMetadata(fetch: FetchNode | null | undefined, enabled = true): AdvisorMetadata | undefined {
	const root = fetch?.entity.name;
	const names = fetch ? advisorEntityNames(fetch) : [];
	const key = names.join(",");
	const [state, setState] = useState<{ key: string; metadata: AdvisorMetadata } | undefined>();

	useEffect(() => {
		if (!enabled || !root || !key || !isDataverseAvailable()) return;
		let cancelled = false;
		const entities = key.split(",");
		Promise.all(
			entities.map(async (entity): Promise<[string, AdvisorEntityInfo]> => {
				const [attributes, lengths, metadata] = await Promise.allSettled([
					loadAllEntityAttributes(entity),
					loadStringMaxLengths(entity),
					entity === root ? loadEntityMetadata(entity) : Promise.resolve(undefined),
				]);
				const info: AdvisorEntityInfo = {
					primaryIdAttribute: metadata.status === "fulfilled" ? metadata.value?.PrimaryIdAttribute : undefined,
				};
				if (attributes.status === "fulfilled") {
					const maxLengths = lengths.status === "fulfilled" ? lengths.value : undefined;
					info.attributes = new Map(
						attributes.value.map((a): [string, AdvisorAttributeInfo] => [
							a.LogicalName,
							{
								typeName: toAttributeTypeInfo(a).typeName,
								sourceType: typeof a.SourceType === "number" ? a.SourceType : undefined,
								maxLength: maxLengths?.get(a.LogicalName),
							},
						]),
					);
				}
				return [entity, info];
			}),
		).then((entries) => {
			if (!cancelled) setState({ key: `${root}|${key}`, metadata: new Map(entries) });
		});
		return () => {
			cancelled = true;
		};
	}, [root, key, enabled]);

	return state && state.key === `${root}|${key}` ? state.metadata : undefined;
}

export interface DataverseCheckResult {
	/** What was checked, compared with the current FetchXML to tell when the result is stale */
	key: string;
	checkedAt: Date;
	check: DataverseCheck;
}

export interface DataverseCheckState {
	checking: boolean;
	result?: DataverseCheckResult;
	error?: string;
}

/**
 * "Check with Dataverse": runs ValidateFetchXmlExpression on demand only (never automatically).
 * The result stays until the next check; the caller compares its key with the current FetchXML to
 * show that the query changed since. A new connection remounts the app (AppShell, #66), which
 * forgets it.
 * https://learn.microsoft.com/power-apps/developer/data-platform/webapi/reference/validatefetchxmlexpression
 */
export function useDataverseCheck() {
	const [state, setState] = useState<DataverseCheckState>({ checking: false });
	const requestRef = useRef(0);

	const run = useCallback(async (fetchXml: string, key: string) => {
		const request = ++requestRef.current;
		setState((prev) => ({ ...prev, checking: true, error: undefined }));
		try {
			const response = await validateFetchXmlExpression(fetchXml);
			if (request !== requestRef.current) return;
			setState({
				checking: false,
				result: {
					key,
					checkedAt: new Date(),
					check: {
						issues: response.ValidationResults?.Messages ?? [],
						helplink: response.ValidationResults?.Helplink,
					},
				},
			});
		} catch (error) {
			if (request !== requestRef.current) return;
			// Dataverse's own message, made actionable where the code is documented (#96)
			setState((prev) => ({ ...prev, checking: false, error: formatDataverseError(error, "Couldn't check with Dataverse.") }));
		}
	}, []);

	const dismissError = useCallback(() => setState((prev) => ({ ...prev, error: undefined })), []);

	return { ...state, run, dismissError };
}
