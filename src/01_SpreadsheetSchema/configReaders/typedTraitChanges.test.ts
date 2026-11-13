import { describe, expect, it } from "vitest";

import type { ColumnConfigStored, ValueConfigsBase } from "./makeConfigs";
import { typedTraitChanges } from "./typedTraitChanges";

function statusColumn(valueName: string): ColumnConfigStored {
  return {
    columnId: "c:led:status",
    header: "Status",
    valueName,
    isFormula: false,
    emptyValueAllowed: false,
    customDefaultValue: null,
    formula: null,
    columnType: { type: "DROPDOWN", rule: null },
    actionRowCell: null,
    groupHeading1: null,
  } as ColumnConfigStored;
}

function configs(valueConfigs: ValueConfigsBase) {
  return {
    columnConfigs: { ledger: { status: statusColumn("listStatus") } },
    valueConfigs,
  };
}

describe("typedTraitChanges", () => {
  it("lists a dropdown's options changed", () => {
    const changes = typedTraitChanges({
      previous: configs({ listStatus: ["Paid", "Unpaid"] }),
      regenerated: configs({ listStatus: ["Paid", "Late"] }),
      sent: [],
    });

    expect(changes).toEqual([
      {
        table: "ledger",
        column: "status",
        trait: "options",
        from: "Paid, Unpaid",
        to: "Paid, Late",
      },
    ]);
  });
});
