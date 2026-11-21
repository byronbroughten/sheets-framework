import type { StrictOmit } from "@byronbroughten/utils/obj";

import type { GridRangeProps } from "../../00_Source/RawSource/RawSource.ts";

export type SheetGridRangeProps = StrictOmit<GridRangeProps, "sheetId">;
