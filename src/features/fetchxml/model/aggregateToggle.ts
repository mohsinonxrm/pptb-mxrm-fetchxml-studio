/**
 * Turning "Aggregate query" off (#90b, D4). Settings that only work in aggregate queries are
 * rejected in a normal query (groupby: 0x8004112e, aggregate: 0x8004111a, order alias: 0x80041132),
 * so the editor offers to clean them up. This module only plans the changes; the editor shows them
 * for confirmation and applies them as node updates.
 */
import type { EntityNode, FetchNode, LinkEntityNode, NodeId, OrderNode } from "./nodes";

export interface AggregateOffChange {
	element: "fetch" | "attribute" | "order";
	message: string;
}

export interface AggregateOffPlan {
	removed: AggregateOffChange[];
	converted: AggregateOffChange[];
	/** Node updates to apply (undefined clears a field) */
	updates: { nodeId: NodeId; updates: Record<string, unknown> }[];
	/** Orders that sort by an aggregate value or an unknown alias: removed */
	removeNodeIds: NodeId[];
}

type Container = EntityNode | LinkEntityNode;

/** Attributes an aggregate order alias refers to (same scoping Dataverse uses). */
function aliasScope(order: OrderNode, container: Container, isRoot: boolean, links: LinkEntityNode[]): Container {
	if (isRoot && order.entityname) {
		return links.find((l) => l.alias === order.entityname) ?? links.find((l) => l.name === order.entityname) ?? container;
	}
	return container;
}

export function planAggregateOff(fetch: FetchNode): AggregateOffPlan {
	const plan: AggregateOffPlan = { removed: [], converted: [], updates: [], removeNodeIds: [] };
	const containers: { container: Container; isRoot: boolean }[] = [];
	const visit = (c: Container, isRoot: boolean) => {
		containers.push({ container: c, isRoot });
		c.links.forEach((l) => visit(l, false));
	};
	visit(fetch.entity, true);
	const links = containers.map((c) => c.container).filter((c): c is LinkEntityNode => c.type === "link-entity");

	for (const { container } of containers) {
		for (const attr of container.attributes) {
			const parts = [
				attr.groupby && "Group by",
				attr.aggregate && `aggregate="${attr.aggregate}"`,
				attr.dategrouping && `dategrouping="${attr.dategrouping}"`,
				attr.usertimezone !== undefined && "time zone",
				attr.distinct && "unique values",
			].filter((p): p is string => typeof p === "string");
			if (parts.length === 0) continue;
			plan.removed.push({
				element: "attribute",
				message: `${attr.name}${attr.alias ? ` (${attr.alias})` : ""}: remove ${parts.join(", ")}`,
			});
			plan.updates.push({
				nodeId: attr.id,
				updates: { groupby: undefined, aggregate: undefined, dategrouping: undefined, usertimezone: undefined, distinct: undefined },
			});
		}
	}

	const removedOrders: AggregateOffChange[] = [];
	for (const { container, isRoot } of containers) {
		for (const order of container.orders) {
			if (!order.alias) continue;
			const source = aliasScope(order, container, isRoot, links).attributes.find((a) => a.alias === order.alias);
			if (source?.groupby) {
				plan.converted.push({ element: "order", message: `Sort by alias ${order.alias} → sort by attribute ${source.name}` });
				plan.updates.push({ nodeId: order.id, updates: { alias: undefined, attribute: source.name } });
			} else {
				removedOrders.push({
					element: "order",
					message: source
						? `Sort by ${order.alias} (it sorts by an aggregate value)`
						: `Sort by ${order.alias} (the alias isn't defined on that entity)`,
				});
				plan.removeNodeIds.push(order.id);
			}
		}
	}
	plan.removed.push(...removedOrders);

	if (fetch.options.aggregateLimit !== undefined) {
		plan.removed.push({ element: "fetch", message: "Aggregate row limit (aggregatelimit)" });
	}
	return plan;
}

export function hasAggregateOffChanges(plan: AggregateOffPlan): boolean {
	return plan.removed.length > 0 || plan.converted.length > 0;
}
