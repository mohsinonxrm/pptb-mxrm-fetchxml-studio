import { describe, expect, it } from "vitest";
import { odataString } from "./odata";

describe("odataString", () => {
	it("leaves ordinary names alone", () => {
		expect(odataString("account")).toBe("account");
		expect(odataString("new_custom_entity")).toBe("new_custom_entity");
	});

	it("doubles single quotes so the value can't end the literal", () => {
		expect(odataString("o'brien")).toBe("o''brien");
		expect(odataString("x' or 1 eq 1 or 'y")).toBe("x'' or 1 eq 1 or ''y");
	});
});
