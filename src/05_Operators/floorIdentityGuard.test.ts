import { describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import type {
  ColumnConfigStored,
  TableConfigsBase,
  TableConfigStored,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  assertFloorIdentityUnchanged,
  type FloorIdentitySource,
} from "./floorIdentityGuard";

const { columnConfigs } = installedConfigs();

const tableIdHeader = columnConfigs.tableConfig.tableId.header;

function column(columnId: string): ColumnConfigStored {
  return { ...columnConfigs.tableConfig.tableId, columnId };
}

const tableConfigEntry: TableConfigStored = {
  tableId: "tbl-floor0001",
  tableName: "tableConfig",
  sheetGid: 1,
  idPrefix: "scf",
  hasIdColumn: false,
  hasNameColumn: false,
};

function source(props: {
  tableConfigs?: TableConfigsBase;
  tableConfigColumnId?: string;
}): FloorIdentitySource {
  return {
    tableConfigs: props.tableConfigs ?? { tableConfig: tableConfigEntry },
    columnConfigs:
      props.tableConfigColumnId === undefined
        ? {}
        : { tableConfig: { tableId: column(props.tableConfigColumnId) } },
  };
}

describe("assertFloorIdentityUnchanged", () => {
  it("throws naming the sheet, the header, and the previous and new column ID of a floor column", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({ tableConfigColumnId: "c:scf:aaa" }),
        next: source({ tableConfigColumnId: "c:scf:bbb" }),
      }),
    ).toThrow(
      `Floor column "${tableIdHeader}" on "tableConfig" had column ID "c:scf:aaa" and is now "c:scf:bbb".`,
    );
  });

  it("throws naming the floor tab and its previous and new Table ID", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({}),
        next: source({
          tableConfigs: {
            tableConfig: { ...tableConfigEntry, tableId: "tbl-floor0002" },
          },
        }),
      }),
    ).toThrow(
      'Floor tab "tableConfig" Table ID was "tbl-floor0001" and is now "tbl-floor0002".',
    );
  });

  it("throws naming the floor tab and its previous and new GID", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({}),
        next: source({
          tableConfigs: { tableConfig: { ...tableConfigEntry, sheetGid: 2 } },
        }),
      }),
    ).toThrow('Floor tab "tableConfig" GID was 1 and is now 2.');
  });

  it("ignores a changed Value Config, which is no longer floor", () => {
    const valueConfigEntry: TableConfigStored = {
      ...tableConfigEntry,
      tableId: "tbl-value0001",
      tableName: "valueConfig",
      idPrefix: "vcf",
    };
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({
          tableConfigs: {
            tableConfig: tableConfigEntry,
            valueConfig: valueConfigEntry,
          },
        }),
        next: source({
          tableConfigs: {
            tableConfig: tableConfigEntry,
            valueConfig: {
              ...valueConfigEntry,
              tableId: "tbl-value0002",
              sheetGid: 2,
              idPrefix: "zzz",
            },
          },
        }),
      }),
    ).not.toThrow();
  });

  it("passes when the floor identities match", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({ tableConfigColumnId: "c:scf:aaa" }),
        next: source({ tableConfigColumnId: "c:scf:aaa" }),
      }),
    ).not.toThrow();
  });

  it("skips a floor tab absent from the previous table configs", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({ tableConfigs: {} }),
        next: source({}),
      }),
    ).not.toThrow();
  });

  it("skips a floor column absent from the previous column configs", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({}),
        next: source({ tableConfigColumnId: "c:scf:bbb" }),
      }),
    ).not.toThrow();
  });
});
