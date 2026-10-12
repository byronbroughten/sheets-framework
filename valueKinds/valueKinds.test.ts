import { describe, expect, it } from "vitest";

import type {
  ColumnValueDeclared,
  ColumnValueName,
} from "../src/01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { VnToCvn } from "../src/01_SpreadsheetSchema/configReaders/valueSchemas";
import { SpreadsheetNamed } from "../src/04_SpreadsheetNamed/SpreadsheetNamed";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../src/testSupport/fakeSheetsService";
import {
  type FakeBodyRow,
  fakeTableSheet,
} from "../src/testSupport/fakeSheetsService/fakeTableSheet";
import { assertType, type IsExactly } from "../src/testSupport/typeAssertions";
import { appConfigs } from "./appConfigs";

const widgetGid = appConfigs.tableConfigs.widget.sheetGid;
const widgetColumnNames = ["status", "itemName", "optionalItemName"] as const;
type WidgetColumnName = (typeof widgetColumnNames)[number];
const topBodyRowIndex = 4;

function fetchedWidget(bodyRows: FakeBodyRow<WidgetColumnName>[] = [{}]) {
  const { grid } = stubSheetsService({
    sheets: [
      fakeTableSheet.build({
        sheetId: widgetGid,
        title: "Widget",
        columnConfigs: appConfigs.columnConfigs.widget,
        columnNames: widgetColumnNames,
        bodyRows,
      }),
    ],
  });
  const ss = SpreadsheetNamed.init();
  const widget = ss.table("widget");
  widget.prepFetchColumnsFull(...widgetColumnNames);
  ss.fetchAllPrepped();
  return { ss, widget, grid };
}

function topBodyRow(grid: FakeSheetsService["grid"]) {
  return grid.sheet(widgetGid).values({
    startRowIndex: topBodyRowIndex,
    endRowIndex: topBodyRowIndex + 1,
    endColumnIndex: widgetColumnNames.length,
  })[0];
}

describe("a fixed-list column", () => {
  it("defaults to its list's first member", () => {
    const { ss, widget, grid } = fetchedWidget();
    widget.appendRowWithVals({ itemName: "Item A" });
    ss.batchUpdateGSheets();

    expect(topBodyRow(grid)?.[0]).toBe("Open");
  });

  it("rejects a value outside its list", () => {
    const { widget } = fetchedWidget([{ status: "Open" }]);
    const status = widget.row(0).cell("status");

    expect(() => status.setValueType("widgetStatus", "Archived")).toThrow(
      /widgetStatus/,
    );
  });

  it("writes a member that isn't a string", () => {
    const { ss, widget, grid } = fetchedWidget([{ status: "Open" }]);
    widget.row(0).cell("status").setValueType("widgetStatus", 3);
    ss.batchUpdateGSheets();

    expect(topBodyRow(grid)?.[0]).toBe(3);
  });

  it("types its value as the list's literal union, number members included", () => {
    assertType<
      IsExactly<ColumnValueDeclared<"widget", "status">, "Open" | "Closed" | 3>
    >(true);
    assertType<
      IsExactly<
        VnToCvn<ColumnValueName<"widget", "status">>,
        "string" | "number"
      >
    >(true);
  });
});

describe("a live-reference column", () => {
  it("rejects a value that isn't a string", () => {
    const { widget } = fetchedWidget([{ itemName: "Item A" }]);

    expect(() =>
      widget.row(0).cell("itemName").setValueType("itemName", 3),
    ).toThrow(/string/i);
  });

  it("accepts a string no list declares", () => {
    const { ss, widget, grid } = fetchedWidget([{ itemName: "Item A" }]);
    widget
      .row(0)
      .cell("itemName")
      .setValueType("itemName", "A row added after the build");
    ss.batchUpdateGSheets();

    expect(topBodyRow(grid)?.[1]).toBe("A row added after the build");
  });

  it("defaults to blank, which Empty value allowed then accepts or rejects", () => {
    const { ss, widget, grid } = fetchedWidget();
    const row = widget.appendRowWithVals({});
    ss.batchUpdateGSheets();

    expect(topBodyRow(grid)?.slice(1)).toEqual(["", ""]);
    expect(row.value("optionalItemName")).toBe("");
    expect(() => row.value("itemName")).toThrow(/itemName/);
  });

  it("types its value as string", () => {
    assertType<IsExactly<ColumnValueDeclared<"widget", "itemName">, string>>(
      true,
    );
    assertType<
      IsExactly<VnToCvn<ColumnValueName<"widget", "itemName">>, "string">
    >(true);
  });
});
