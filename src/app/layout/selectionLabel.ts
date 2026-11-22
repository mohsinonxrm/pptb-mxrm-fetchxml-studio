/**
 * What the builder pane's selection bar says about the selected node (#234): its kind, then what
 * names it ("Condition · revenue"), in the words the tree and the Properties panel use.
 */

/** The fields a node's label reads; every tree node type fits */
export interface LabelledNode {
	type: string;
	name?: string;
	alias?: string;
	attribute?: string;
	conjunction?: "and" | "or";
}

/** "Kind · name" for the selection bar, or just the kind when nothing names the node */
export function selectionLabel(node: LabelledNode | null | undefined): string {
	if (!node) return "Nothing selected";
	const withName = (kind: string, name: string | undefined) =>
		name ? `${kind} · ${name}` : kind;
	switch (node.type) {
		case "fetch":
			return "Query options";
		case "entity":
			return withName("Entity", node.name);
		case "all-attributes":
			return "All attributes";
		case "attribute":
			return withName("Attribute", node.alias || node.name);
		case "order":
			return withName("Sort", node.alias || node.attribute);
		case "filter":
			return node.conjunction === "or" ? "Filter (or)" : "Filter (and)";
		case "condition":
			return withName("Condition", node.attribute);
		case "link-entity":
			return withName("Link-entity", node.alias ? `${node.name} (${node.alias})` : node.name);
		default:
			return "Node";
	}
}
