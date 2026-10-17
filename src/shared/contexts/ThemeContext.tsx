/**
 * Theme context for accessing dark/light mode throughout the app
 */

import type { ReactNode } from "react";
import { ThemeContext } from "./useTheme";

export function ThemeProvider({ isDark, children }: { isDark: boolean; children: ReactNode }) {
	return <ThemeContext.Provider value={{ isDark }}>{children}</ThemeContext.Provider>;
}
