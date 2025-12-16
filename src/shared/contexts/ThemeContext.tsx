/**
 * Theme context for accessing dark/light mode throughout the app
 */

import { createContext, useContext, type ReactNode } from "react";

interface ThemeContextValue {
	isDark: boolean;
}

const ThemeContext = createContext<ThemeContextValue>({ isDark: false });

export function ThemeProvider({ isDark, children }: { isDark: boolean; children: ReactNode }) {
	return <ThemeContext.Provider value={{ isDark }}>{children}</ThemeContext.Provider>;
}

export function useTheme(): ThemeContextValue {
	return useContext(ThemeContext);
}
