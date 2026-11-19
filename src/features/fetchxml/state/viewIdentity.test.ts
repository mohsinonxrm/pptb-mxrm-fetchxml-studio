/**
 * The view the query is (#278): set when a view loads and when it's saved, kept across edits and
 * across loading FetchXML into the tree (unlike loadedView), and cleared by a new entity or New.
 */
import { act, createElement } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { BuilderProvider } from "./builderStore";
import { useBuilder, type BuilderContextValue } from "./builderContext";
import { generateFetchXml } from "../model/fetchxml";
import { cleanup, render } from "../../../shared/testing/ui";

let current: BuilderContextValue | null = null;
function Probe() {
	current = useBuilder();
	return null;
}
const b = () => {
	if (!current) throw new Error("builder not rendered");
	return current;
};
const run = (fn: (builder: BuilderContextValue) => void) =>
	act(async () => {
		fn(b());
	});

const VIEW_XML =
	"<fetch><entity name='account'><attribute name='name'/><order attribute='name'/></entity></fetch>";
const LAYOUT =
	"<grid name='resultset' object='1' jump='name' select='1' icon='1' preview='1'><row name='result' id='accountid'><cell name='name' width='300'/></row></grid>";
const viewInfo = {
	id: "v1",
	type: "personal" as const,
	entitySetName: "accounts",
	name: "My accounts",
};

afterEach(() => {
	cleanup();
	current = null;
});

describe("viewIdentity", () => {
	it("is set by Load view, with the view's FetchXML and LayoutXML as the baseline", async () => {
		await render(createElement(BuilderProvider, null, createElement(Probe)));
		expect(b().viewIdentity).toBeNull();
		await run((x) => void x.loadView(VIEW_XML, viewInfo, LAYOUT));
		expect(b().viewIdentity).toEqual({
			name: "My accounts",
			baselineFetchXml: b().loadedView?.originalFetchXml,
			baselineLayoutXml: LAYOUT,
		});
	});

	it("stays through edits and loading FetchXML into the tree, where loadedView goes", async () => {
		await render(createElement(BuilderProvider, null, createElement(Probe)));
		await run((x) => void x.loadView(VIEW_XML, viewInfo, LAYOUT));
		const identity = b().viewIdentity;
		await run((x) => x.addOrder(x.fetchQuery!.entity.id, "createdon"));
		expect(b().loadedView).toBeNull();
		expect(b().viewIdentity).toEqual(identity);
		await run((x) => void x.loadFetchXml(VIEW_XML));
		expect(b().viewIdentity).toEqual(identity);
	});

	it("is what was saved after a save", async () => {
		await render(createElement(BuilderProvider, null, createElement(Probe)));
		await run((x) => x.setEntity("account"));
		const saved = generateFetchXml(b().fetchQuery!);
		await run((x) =>
			x.setLoadedView({ ...viewInfo, name: "Saved view", originalFetchXml: saved }, LAYOUT),
		);
		expect(b().viewIdentity).toEqual({
			name: "Saved view",
			baselineFetchXml: saved,
			baselineLayoutXml: LAYOUT,
		});
	});

	it("goes with a new entity and with New", async () => {
		await render(createElement(BuilderProvider, null, createElement(Probe)));
		await run((x) => void x.loadView(VIEW_XML, viewInfo, LAYOUT));
		await run((x) => x.setEntity("contact"));
		expect(b().viewIdentity).toBeNull();
		await run((x) => void x.loadView(VIEW_XML, viewInfo, LAYOUT));
		await run((x) => x.newQuery());
		expect(b().viewIdentity).toBeNull();
	});
});
