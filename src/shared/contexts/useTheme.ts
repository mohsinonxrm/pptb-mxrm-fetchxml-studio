/**
 * Theme context value and hook, kept separate from ThemeProvider so the
 * provider module only exports components (React Fast Refresh).
 */

import { createContext, useContext } from "react";

export interface ThemeContextValue {
	isDark: boolean;
}

export const ThemeContext = createContext<ThemeContextValue>({ isDark: false });

export function useTheme(): ThemeContextValue {
	return useContext(ThemeContext);
}
