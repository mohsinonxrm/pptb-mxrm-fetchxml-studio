/**
 * Metadata for the property editors (#90b): attribute types drive which controls and values are
 * offered (attributeCapabilities.ts); entity metadata drives primary key, retention and hierarchy.
 * Everything comes from the cached metadata loaders; unknown metadata never blocks anything.
 */
import { useEffect, useState } from "react";
import type { AttributeMetadata, EntityMetadata } from "../../../api/pptbClient";
import { loadAllEntityAttributes, loadEntityMetadata, loadIsHierarchical } from "../../../api/dataverseMetadata";
import { toAttributeTypeInfo, type AttributeTypeInfo } from "../../../model/attributeCapabilities";

/**
 * Every attribute of an entity, for looking up attributes already in the query (not a picker, so
 * the "Advanced Find only" setting doesn't apply, #57); undefined while loading / unavailable.
 */
export function useEntityAttributes(entityName: string | undefined): AttributeMetadata[] | undefined {
	const [state, setState] = useState<{ entity?: string; attributes?: AttributeMetadata[] }>({});
	useEffect(() => {
		if (!entityName) return;
		let cancelled = false;
		loadAllEntityAttributes(entityName)
			.then((attributes) => !cancelled && setState({ entity: entityName, attributes }))
			.catch(() => !cancelled && setState({ entity: entityName, attributes: undefined }));
		return () => {
			cancelled = true;
		};
	}, [entityName]);
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
		});
		return () => {
			cancelled = true;
		};
	}, [entityName]);
	return state.entity === entityName ? state.info : {};
}
