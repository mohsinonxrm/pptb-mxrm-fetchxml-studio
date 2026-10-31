/**
 * Metadata for the property editors (#90b): attribute types drive which controls and values are
 * offered (attributeCapabilities.ts); entity metadata drives primary key, retention and hierarchy.
 * Everything comes from the cached metadata loaders; unknown metadata never blocks anything. A
 * failed load is recorded so the editor says why and offers Retry (#201, MetadataLoadWarning).
 */
import { useEffect, useState } from "react";
import type { AttributeMetadata, EntityMetadata } from "../../../api/pptbClient";
import { loadAllEntityAttributes, loadEntityMetadata, loadIsHierarchical } from "../../../api/dataverseMetadata";
import { toAttributeTypeInfo, type AttributeTypeInfo } from "../../../model/attributeCapabilities";
import { recordMetadataLoad, useMetadataRetryGeneration } from "./metadataLoadErrors";

/**
 * Every attribute of an entity, for looking up attributes already in the query (not a picker, so
 * the "Advanced Find only" setting doesn't apply, #57); undefined while loading / unavailable.
 */
export function useEntityAttributes(entityName: string | undefined): AttributeMetadata[] | undefined {
	const [state, setState] = useState<{ entity?: string; attributes?: AttributeMetadata[] }>({});
	const retry = useMetadataRetryGeneration();
	useEffect(() => {
		if (!entityName) return;
		let cancelled = false;
		loadAllEntityAttributes(entityName)
			.then((attributes) => {
				if (cancelled) return;
				setState({ entity: entityName, attributes });
				recordMetadataLoad("columns", entityName);
			})
			.catch((error: unknown) => {
				if (cancelled) return;
				setState({ entity: entityName, attributes: undefined });
				recordMetadataLoad("columns", entityName, error);
			});
		return () => {
			cancelled = true;
		};
	}, [entityName, retry]);
	return state.entity === entityName ? state.attributes : undefined;
}

/** Type info for one attribute (undefined when unknown: still loading, typed by hand, or failed). */
export function useAttributeTypeInfo(entityName: string | undefined, attributeName: string | undefined): AttributeTypeInfo | undefined {
	const attributes = useEntityAttributes(entityName);
	const attribute = attributeName ? attributes?.find((a) => a.LogicalName === attributeName) : undefined;
	return attribute ? toAttributeTypeInfo(attribute) : undefined;
}

export interface EntityInfo {
	metadata?: EntityMetadata;
	/** Has a hierarchical relationship; undefined while unknown */
	isHierarchical?: boolean;
}

export function useEntityInfo(entityName: string | undefined): EntityInfo {
	const [state, setState] = useState<{ entity?: string; info: EntityInfo }>({ info: {} });
	const retry = useMetadataRetryGeneration();
	useEffect(() => {
		if (!entityName) return;
		let cancelled = false;
		Promise.allSettled([loadEntityMetadata(entityName), loadIsHierarchical(entityName)]).then(([m, h]) => {
			if (cancelled) return;
			setState({
				entity: entityName,
				info: {
					metadata: m.status === "fulfilled" ? m.value : undefined,
					isHierarchical: h.status === "fulfilled" ? h.value : undefined,
				},
			});
			const failed = [m, h].find((r): r is PromiseRejectedResult => r.status === "rejected");
			recordMetadataLoad("details", entityName, failed?.reason);
		});
		return () => {
			cancelled = true;
		};
	}, [entityName, retry]);
	return state.entity === entityName ? state.info : {};
}
