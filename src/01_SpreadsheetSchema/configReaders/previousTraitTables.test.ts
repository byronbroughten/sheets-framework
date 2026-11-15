import { describe, expect, it } from "vitest";

import type { ColumnTypeBaseline } from "../columnBaseline";
import type { BraceConfigs } from "./BraceFormula";
import { previousTraitTables } from "./previousTraitTables";
import { SyncPlanner, type TraitReadings } from "./SyncPlanner";

function column(
  columnId: string,
  header: string,
  columnType?: ColumnTypeBaseline,
) {
  return {
    columnId,
    header,
    valueName: "string",
    isFormula: false,
    emptyValueAllowed: true,
    customDefaultValue: null,
    columnType,
  };
}

function table(tableId: string, tableName: string, sheetGid: number) {
  return {
    tableId,
    tableName,
    sheetGid,
    idPrefix: tableId,
    hasIdColumn: true,
    hasNameColumn: false,
  };
}

function rangeDropdown(value: string): ColumnTypeBaseline {
  return {
    type: "DROPDOWN",
    rule: { condition: "ONE_OF_RANGE", values: [value] },
  };
}

const listDropdown: ColumnTypeBaseline = {
  type: "DROPDOWN",
  rule: { condition: "ONE_OF_LIST", values: ["High", "Size[cm]"] },
};

// Keyed by installed Table keys, so an alteration on them type-checks; the item Table's name rule names the log Table's Rent column.
function configsWithNameRule(nameType: ColumnTypeBaseline): BraceConfigs {
  return {
    tableConfigs: {
      log: table("log", "Unit_list", 1),
      item: table("item", "Items", 2),
    },
    columnConfigs: {
      log: { rent: column("c:log:rent", "Rent") },
      item: {
        name: column("c:itm:name", "Name", nameType),
        priority: column("c:itm:priority", "Priority", listDropdown),
      },
    },
  };
}

function itemColumnTypes(configs: BraceConfigs): Record<string, unknown> {
  const columns = previousTraitTables(configs).item?.columns ?? {};
  return Object.fromEntries(
    Object.entries(columns).map(([columnName, reading]) => [
      columnName,
      reading?.columnType,
    ]),
  );
}

describe("previousTraitTables", () => {
  it("records a list rule and a rule in brace form", () => {
    const braced = rangeDropdown("=${log[rent]}");

    expect(itemColumnTypes(configsWithNameRule(braced))).toEqual({
      name: braced,
      priority: listDropdown,
    });
  });

  it("records no column type for a rule written before rules were braced", () => {
    const oldForm = configsWithNameRule(rangeDropdown("=Unit_list[Rent]"));

    expect(itemColumnTypes(oldForm)).toEqual({
      name: undefined,
      priority: listDropdown,
    });
  });

  it("takes live's rule as the baseline for an old-form rule, so an alteration sends without a conflict", () => {
    const previous: TraitReadings = {
      tables: previousTraitTables(
        configsWithNameRule(rangeDropdown("=Unit_list[Rent]")),
      ),
      tabs: {},
    };
    const live: TraitReadings = {
      tables: {
        ...previous.tables,
        item: {
          traits: {},
          columns: {
            name: {
              header: "Name",
              columnType: rangeDropdown("=${log[rent]}"),
            },
          },
        },
      },
      tabs: {},
    };
    const intended = rangeDropdown("=${log[rent]}*2");

    const result = SyncPlanner.init({
      alterations: {
        tables: { item: { columns: { name: { columnType: intended } } } },
      },
      tabsByTable: {},
      forcedTraits: [],
    }).plan({ previous, live });

    expect(result.stops).toEqual([]);
    expect(result.plan).toHaveLength(1);
  });
});
