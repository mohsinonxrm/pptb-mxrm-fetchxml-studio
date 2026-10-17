/** Option value for the "None" choice in optional dropdowns. */
export const NONE_OPTION = "none";

/**
 * Value to store for an optional dropdown: "None" (or nothing) removes the attribute instead of
 * writing the literal "none", which Dataverse rejects (#88).
 */
export function optionalValue(optionValue: string | undefined): string | undefined {
	return !optionValue || optionValue === NONE_OPTION ? undefined : optionValue;
}
