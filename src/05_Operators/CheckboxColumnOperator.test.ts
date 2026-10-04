import { beforeEach, describe, expect, it } from "vitest";

import {
  type ColumnValueName,
  getColumnTraitByName,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { ColumnNamedProps } from "../04_SpreadsheetNamed/ClassBases/ColumnBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import {
  type CheckboxColumnName,
  CheckboxColumnOperator,
} from "./CheckboxColumnOperator";

const runItemGid = getSheetTraitByName("runItem", "sheetGid");
const columnIds = [
  getColumnTraitByName("runItem", "id", "columnId"),
  getColumnTraitByName("runItem", "selected", "columnId"),
];
const selectColIndex = 1;
const endRowIndex = 7;

function seedSelectedColumn(selectCells: readonly FakeCell[]) {
  return stubSheetsService({
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({
          0: columnIds,
          3: ["ID", "Selected"],
          4: ["r:rit:row4", selectCells[0] ?? null],
          5: ["r:rit:row5", selectCells[1] ?? null],
          6: ["r:rit:row6", selectCells[2] ?? null],
        }),
        table: { endRowIndex },
      },
    ],
  });
}

function selectedProps(): ColumnNamedProps<"runItem", "selected"> {
  return {
    ...SpreadsheetNamed.initSpreadsheetNamedProps(),
    sheetName: "runItem",
    columnName: "selected",
  };
}

function initOperatorWithFetchedColumn() {
  const props = selectedProps();
  const operator = new CheckboxColumnOperator(props);
  operator.column.prepFetchFull();
  new SpreadsheetNamed(props).fetchAllPrepped();
  return operator;
}

const selectedColumnRange = {
  startRowIndex: 4,
  endRowIndex,
  startColumnIndex: selectColIndex,
  endColumnIndex: selectColIndex + 1,
};

beforeEach(() => {
  stubLogger();
});

describe("CheckboxColumnOperator.rowIndexesChecked", () => {
  it("returns only the row indexes whose checkbox is checked", () => {
    seedSelectedColumn([true, false, true]);
    expect(initOperatorWithFetchedColumn().rowIndexesChecked).toEqual([0, 2]);
  });

  it("counts an empty cell as unchecked", () => {
    seedSelectedColumn([null, null, true]);
    expect(initOperatorWithFetchedColumn().rowIndexesChecked).toEqual([2]);
  });
});

describe("CheckboxColumnOperator.uncheckActiveCells", () => {
  it("unticks a pruned sheet's remaining rows and leaves the pruned rows ticked", () => {
    const { grid } = seedSelectedColumn([true, true, true]);
    const operator = initOperatorWithFetchedColumn();

    operator.sheet.raw.removeRowsExcept(0, 2);
    operator.uncheckActiveCells();
    operator.ss.batchUpdateGSheets();

    expect(grid.sheet(runItemGid).values(selectedColumnRange)).toEqual([
      [false],
      [true],
      [false],
    ]);
  });

  it("normalizes an untouched empty cell to an explicit false", () => {
    const { grid } = seedSelectedColumn([true, null, null]);
    const operator = initOperatorWithFetchedColumn();

    operator.sheet.raw.removeRowsExcept(1, 2);
    operator.uncheckActiveCells();
    operator.ss.batchUpdateGSheets();

    expect(grid.sheet(runItemGid).values(selectedColumnRange)).toEqual([
      [true],
      [false],
      [false],
    ]);
  });
});

describe("CheckboxColumnOperator, column constraint", () => {
  it("accepts a declared non-formula checkbox column and rejects anything else", () => {
    const checkbox = new CheckboxColumnOperator({
      ...SpreadsheetNamed.initSpreadsheetNamedProps(),
      sheetName: "runItem",
      columnName: "selected",
    });
    const text = new CheckboxColumnOperator({
      ...SpreadsheetNamed.initSpreadsheetNamedProps(),
      sheetName: "runItem",
      // @ts-expect-error a string column is not a checkbox column
      columnName: "result",
    });
    const sampled = new CheckboxColumnOperator({
      ...SpreadsheetNamed.initSpreadsheetNamedProps(),
      sheetName: "valueTypes",
      // @ts-expect-error an undeclared column that merely holds a boolean is not one either
      columnName: "sampledBoolean",
    });
    expect(checkbox.schema.valueName).toBe("checkbox");
    expect(checkbox.schema.isFormula).toBe(false);
    expect(text.schema.valueName).toBe("string");
    expect(sampled.schema.valueName).toBe("boolean");
    expect(sampled.schema.isFormula).toBe(false);
    assertType<
      IsExactly<
        ColumnValueName<"runItem", CheckboxColumnName<"runItem">>,
        "checkbox"
      >
    >(true);
  });

  // A config-describing sheet, so regeneration can't churn the expected union.
  it("names exactly the declared non-formula checkbox columns of a sheet", () => {
    assertType<IsExactly<CheckboxColumnName<"sheetConfig">, "letApiAccess">>(
      true,
    );
  });
});
