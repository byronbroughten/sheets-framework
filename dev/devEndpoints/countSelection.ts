import type { Endpoint } from "../../src/framework";

export const countSelection: Endpoint<"runItem"> = {
  timeLastRan: "startTime",
  runStatus: "runStatus",
  selector: { column: "selected" },
  action: (_ss, { selectedRowIndexes }) =>
    `Counted ${selectedRowIndexes.length} selected row(s)`,
};
