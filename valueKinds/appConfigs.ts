// Without it, a program that reaches this file only by dynamic import rejects the augmentation (TS2664).
import type {} from "@byronbroughten/sheets-framework";

import { columnConfigs as devColumnConfigs } from "../dev/generated/columnConfigs";
import { tableConfigs as devTableConfigs } from "../dev/generated/tableConfigs";
import {
  makeColumnConfigs,
  makeTableConfigs,
  makeValueConfigs,
} from "../src/01_SpreadsheetSchema/configReaders/makeConfigs";

const tableConfigs = makeTableConfigs({
  ...devTableConfigs,
  widget: {
    tableId: "widget",
    tableName: "widget",
    sheetGid: 1200001,
    idPrefix: "wdg",
    hasIdColumn: false,
    hasNameColumn: false,
  },
});

const columnConfigs = makeColumnConfigs({
  ...devColumnConfigs,
  widget: {
    status: {
      columnId: "c:wdg:status",
      header: "Status",
      valueName: "widgetStatus",
      isFormula: false,
      emptyValueAllowed: false,
      customDefaultValue: null,
    },
    itemName: {
      columnId: "c:wdg:itemName",
      header: "Item name",
      valueName: "itemName",
      isFormula: false,
      emptyValueAllowed: false,
      customDefaultValue: null,
    },
    optionalItemName: {
      columnId: "c:wdg:optionalItemName",
      header: "Optional item name",
      valueName: "itemName",
      isFormula: false,
      emptyValueAllowed: true,
      customDefaultValue: null,
    },
  },
});

const valueConfigs = makeValueConfigs({
  widgetStatus: ["Open", "Closed", 3],
  itemName: { kind: "liveReference" },
});

export const appConfigs = {
  tableConfigs,
  columnConfigs,
  valueConfigs,
};

declare module "@byronbroughten/sheets-framework" {
  interface Register {
    configs: typeof appConfigs;
  }
}
