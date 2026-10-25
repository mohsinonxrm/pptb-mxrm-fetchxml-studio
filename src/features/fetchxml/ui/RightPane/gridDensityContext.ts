import { createContext, useContext } from "react";
import type { GridDensity } from "../../model/displaySettings";

/** The results grid's density, for cell renderers whose controls must fit the row (#228) */
export const GridDensityContext = createContext<GridDensity>("medium");

export const useGridDensity = (): GridDensity => useContext(GridDensityContext);
