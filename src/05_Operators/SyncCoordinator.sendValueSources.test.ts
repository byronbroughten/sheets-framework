import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeCell,
  type FakeCellValue,
  type FakeSheetProperties,
  type FakeSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  type FakeBodyRow,
  fakeTableSheet,
} from "../testSupport/fakeSheetsService/fakeTableSheet";
import { type ConfigRegeneration, SyncCoordinator } from "./SyncCoordinator";
import {
  columnsByHeader,
  runWithNoAlterations,
  seedFixture,
  sendWithNoAlterations,
} from "./syncCoordinatorTestSupport";

type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator send with no alterations value sources", () => {
  const listGid = 889001;
  const ledgerGid = 889002;
  // The previous configs grant API access, so an API-access fixture takes a Table ID they hold.
  const listTableId = getTableTraitByName("log", "tableId");
  const ledgerTableId = getTableTraitByName("runItem", "tableId");
  const listPaymentTableId = getTableTraitByName("computed", "tableId");

  interface FixtureColumn {
    header: string;
    values?: readonly FakeCell[];
    columnType?: string;
    validationValues?: string[];
    validationConditionType?: string;
    actionRowCell?: FakeCellValue;
    groupHeading1?: string;
  }

  // Omit tableId for a Table the previous configs don't hold.
  function tableSheet(props: {
    sheetId: number;
    title: string;
    tableId?: string;
    columns: readonly FixtureColumn[];
  }): FakeSheetProperties {
    const { columns } = props;
    const headers = columns.map((column) => column.header);
    const rowCount = Math.max(
      1,
      ...columns.map((column) => column.values?.length ?? 0),
    );
    const bodyRows: FakeBodyRow<string>[] = Array.from(
      { length: rowCount },
      (_, rowOffset) =>
        Object.fromEntries(
          columns.map((column) => [
            column.header,
            column.values?.[rowOffset] ?? null,
          ]),
        ),
    );
    return fakeTableSheet.build({
      sheetId: props.sheetId,
      title: props.title,
      tableId: props.tableId,
      name: props.title,
      columnConfigs: columnsByHeader(headers),
      columnNames: headers,
      bodyRows,
      headRows: {
        1: columns.map((column) => column.actionRowCell ?? null),
        2: columns.map((column) => column.groupHeading1 ?? null),
      },
      columnTypes: byColIndex(columns, (column) => column.columnType),
      columnValidationValues: byColIndex(
        columns,
        (column) => column.validationValues,
      ),
      columnValidationConditionTypes: byColIndex(
        columns,
        (column) => column.validationConditionType,
      ),
    });
  }

  function byColIndex<V>(
    columns: readonly FixtureColumn[],
    read: (column: FixtureColumn) => V | undefined,
  ): Record<number, V> {
    return columns.reduce<Record<number, V>>((byIndex, column, colIndex) => {
      const value = read(column);
      if (value !== undefined) byIndex[colIndex] = value;
      return byIndex;
    }, {});
  }

  function ledgerSheet(columns: readonly FixtureColumn[]): FakeSheetProperties {
    return tableSheet({
      sheetId: ledgerGid,
      title: "Ledger",
      tableId: ledgerTableId,
      columns,
    });
  }

  function listSheet(columns: readonly FixtureColumn[]): FakeSheetProperties {
    return tableSheet({
      sheetId: listGid,
      title: "List",
      tableId: listTableId,
      columns,
    });
  }

  function generate() {
    const coordinator = SyncCoordinator.init();
    const files = sendWithNoAlterations(coordinator);
    return {
      valueConfigs: emittedValueConfigs(files),
      ledgerValueName: (columnName: string) =>
        coordinator.columnConfigOperator.newColumnConfigs().ledger?.[columnName]
          ?.valueName,
    };
  }

  // A source fetch reads one whole column; row samples span the Table.
  function fullyFetchedListColumnIndexes(service: FakeSheetsService): number[] {
    const indexes = service.getByDataFilterCalls.flatMap((resource) =>
      (
        (resource as { dataFilters?: { gridRange?: GridRange }[] })
          .dataFilters ?? []
      )
        .map(({ gridRange }) => gridRange)
        .filter(
          (range): range is GridRange =>
            range?.sheetId === listGid &&
            range.startRowIndex === 0 &&
            (range.endColumnIndex ?? 0) - (range.startColumnIndex ?? 0) === 1,
        )
        .map((range) => Val.assert(range.startColumnIndex, "start column")),
    );
    return [...new Set(indexes)].sort();
  }

  function emittedValueConfigs(files: ConfigRegeneration): unknown {
    const match = files.valueConfigs.match(/makeValueConfigs\(([\s\S]*)\);/);
    return JSON.parse(Val.assert(match?.[1], "emitted value configs"));
  }

  it("types a dropdown from a Let api access Table column as that column's fixed list", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          { header: "Payment status", values: ["Paid", "Unpaid", "Late"] },
        ]),
        ledgerSheet([
          {
            header: "Status",
            values: ["Paid"],
            validationValues: ["=List[Payment status]"],
          },
        ]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("status")).toBe("listPaymentStatus");
    expect(valueConfigs).toEqual({
      listPaymentStatus: ["Paid", "Unpaid", "Late"],
    });
  });

  function seedNameAndIdSources(): FakeSheetsService {
    return seedFixture({
      extraSheets: [
        listSheet([
          { header: "ID", values: ["r:lst:a1"] },
          { header: "Name", values: ["Unit 2B"], columnType: "DOUBLE" },
          { header: "Kind", values: ["Flat"] },
        ]),
        ledgerSheet([
          { header: "Unit", validationValues: ["=List[Name]"] },
          { header: "Unit row", validationValues: ["=List[ID]"] },
          { header: "Kind", validationValues: ["=List[Kind]"] },
        ]),
      ],
    });
  }

  it("emits a dropdown from a Table's Name or ID column as a live reference, whatever its type", () => {
    seedNameAndIdSources();

    const { valueConfigs, ledgerValueName } = generate();

    expect(
      ["unit", "unitRow", "kind"].map((columnName) =>
        ledgerValueName(columnName),
      ),
    ).toEqual(["listName", "listId", "listKind"]);
    expect(valueConfigs).toEqual({
      listName: { kind: "liveReference" },
      listId: { kind: "liveReference" },
      listKind: ["Flat"],
    });
  });

  it("fetches no Name source column, only the fixed-list ones and the ID column Fill Row IDs reads", () => {
    const service = seedNameAndIdSources();

    generate();

    // Column 0 is the ID column, read after the regen to fill its blanks.
    expect(fullyFetchedListColumnIndexes(service)).toEqual([0, 2]);
  });

  it("keeps a one-word camel-case header's casing in the value name", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "CapEx", values: ["Roof"] }]),
        ledgerSheet([{ header: "CapEx", validationValues: ["=List[CapEx]"] }]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("capEx")).toBe("listCapEx");
    expect(valueConfigs).toEqual({ listCapEx: ["Roof"] });
  });

  it("types every member by its source's declared type, numbers for number and date", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          { header: "Rent", values: [950, 1200], columnType: "DOUBLE" },
          { header: "Due date", values: [45000, 45031], columnType: "DATE" },
          { header: "Unit", values: ["2B", 101], columnType: "TEXT" },
        ]),
        ledgerSheet([
          { header: "Rent", validationValues: ["=List[Rent]"] },
          { header: "Due", validationValues: ["=List[Due date]"] },
          { header: "Unit", validationValues: ["=List[Unit]"] },
        ]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({
      listRent: [950, 1200],
      listDueDate: [45000, 45031],
      listUnit: ["2B", "101"],
    });
  });

  it("fails the config sync when a declared source holds a member of another primitive, naming the source", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          { header: "Rent", values: [950, "TBD"], columnType: "DOUBLE" },
        ]),
        ledgerSheet([{ header: "Rent", validationValues: ["=List[Rent]"] }]),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Value source "List[Rent]" declares number members, but holds "TBD".',
    );
  });

  it("keeps each member's own primitive on an untyped source", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Score", values: ["N/A", 3, true] }]),
        ledgerSheet([{ header: "Score", validationValues: ["=List[Score]"] }]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({ listScore: ["N/A", 3, true] });
  });

  it("drops blanks and collapses duplicates to their first position, keeping sheet order", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          {
            header: "Payment status",
            values: ["Unpaid", null, "Paid", "Unpaid", "", "Late", "Paid"],
          },
        ]),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
        ]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({
      listPaymentStatus: ["Unpaid", "Paid", "Late"],
    });
  });

  it("gives every dropdown on one source the one value name", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Payment status", values: ["Paid", "Unpaid"] }]),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
          {
            header: "Deposit status",
            validationValues: ["=List[Payment status]"],
          },
        ]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("status")).toBe("listPaymentStatus");
    expect(ledgerValueName("depositStatus")).toBe("listPaymentStatus");
    expect(valueConfigs).toEqual({ listPaymentStatus: ["Paid", "Unpaid"] });
  });

  it("leaves a dropdown untyped, and in the summary, unless it names a column of a Let api access Table", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Payment status", values: ["Paid", "Unpaid"] }]),
        ledgerSheet([
          {
            header: "Non-api",
            values: ["Paid"],
            validationValues: ["=Hidden[Payment status]"],
          },
          {
            header: "Unknown header",
            values: ["Paid"],
            validationValues: ["=List[Nope]"],
          },
          {
            header: "A1 range",
            values: ["Paid"],
            validationValues: ["=List!$A$5:$A$9"],
          },
          {
            header: "Item list",
            values: ["Paid"],
            validationValues: ["Paid", "Unpaid"],
          },
        ]),
        tableSheet({
          sheetId: 889003,
          title: "Hidden",
          columns: [{ header: "Payment status", values: ["Paid"] }],
        }),
      ],
    });
    const logger = stubLogger();

    const coordinator = SyncCoordinator.init();
    const { result, files } = runWithNoAlterations(coordinator);

    expect(emittedValueConfigs(Val.assert(files, "emitted files"))).toEqual({});
    const ledger = coordinator.columnConfigOperator.newColumnConfigs().ledger;
    expect(
      ["nonApi", "unknownHeader", "a1Range", "itemList"].map(
        (columnName) => ledger?.[columnName]?.valueName,
      ),
    ).toEqual(["string", "string", "string", "string"]);
    expect(result.configs.notes).toEqual([
      expect.stringContaining("column(s) across"),
    ]);
    expect(logger.log).toHaveBeenCalledWith(
      'Untyped columns on Table "Ledger" (4): Non-api, Unknown header, A1 range, Item list',
    );
  });

  it("fails the config sync when a source's value name is a framework value name, naming both", () => {
    seedFixture({
      extraSheets: [
        tableSheet({
          sheetId: listGid,
          title: "Number",
          tableId: listTableId,
          columns: [{ header: "#", values: [1, 2] }],
        }),
        ledgerSheet([{ header: "Pick", validationValues: ["=Number[#]"] }]),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Value source "Number[#]" takes the value name "number", which is already the framework value name "number".',
    );
  });

  it("fails the config sync when two sources take one value name, naming both", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Payment status", values: ["Paid"] }]),
        tableSheet({
          sheetId: 889004,
          title: "List Payment",
          tableId: listPaymentTableId,
          columns: [{ header: "Status", values: ["Open"] }],
        }),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
          { header: "State", validationValues: ["=List Payment[Status]"] },
        ]),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Value sources "List[Payment status]" and "List Payment[Status]" both take the value name "listPaymentStatus".',
    );
  });

  describe("the baseline each column records", () => {
    function generateLedgerColumns() {
      const coordinator = SyncCoordinator.init();
      const files = sendWithNoAlterations(coordinator);
      return {
        files,
        ledger: Val.assert(
          coordinator.columnConfigOperator.newColumnConfigs().ledger,
          "Ledger column configs",
        ),
      };
    }

    it("records formula text in brace form, naming its own Table's columns bare, another Table's in full, and this row with @", () => {
      seedFixture({
        extraSheets: [
          listSheet([{ header: "Payment status", values: ["Paid"] }]),
          ledgerSheet([
            { header: "Amount", values: [10] },
            {
              header: "Double",
              values: [
                {
                  value: 20,
                  isFormula: true,
                  formula: "=SINGLE(Ledger[Amount])*2",
                },
              ],
            },
            {
              header: "Statuses",
              values: [
                {
                  value: 1,
                  isFormula: true,
                  formula: '=COUNTA(List[Payment status])&"Ledger[Amount]"',
                },
              ],
            },
          ]),
        ],
      });

      const { files, ledger } = generateLedgerColumns();

      expect(ledger.amount?.formula).toBeNull();
      expect(ledger.double?.formula).toBe("=${@amount}*2");
      expect(ledger.statuses?.formula).toBe(
        '=COUNTA(${list[paymentStatus]})&"Ledger[Amount]"',
      );
      expect(files.columnConfigs).toContain('"formula": "=${@amount}*2"');
    });

    it("records the column type, with the rule a dropdown carries", () => {
      seedFixture({
        extraSheets: [
          listSheet([{ header: "Payment status", values: ["Paid"] }]),
          ledgerSheet([
            { header: "Amount", values: [10], columnType: "CURRENCY" },
            {
              header: "Status",
              values: ["Paid"],
              columnType: "DROPDOWN",
              validationConditionType: "ONE_OF_RANGE",
              validationValues: ["=List[Payment status]"],
            },
            { header: "Note", values: ["x"] },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.columnType).toEqual({
        type: "CURRENCY",
        rule: null,
      });
      expect(ledger.status?.columnType).toEqual({
        type: "DROPDOWN",
        rule: { condition: "ONE_OF_RANGE", values: ["=List[Payment status]"] },
      });
      expect(ledger.note?.columnType).toEqual({ type: null, rule: null });
    });

    it("records each action-row cell as empty, a checkbox or a level-2 group heading", () => {
      seedFixture({
        extraSheets: [
          ledgerSheet([
            { header: "Amount", values: [10] },
            { header: "Run", values: [false], actionRowCell: false },
            { header: "Due", values: [1], actionRowCell: "Dates" },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.actionRowCell).toBeNull();
      expect(ledger.run?.actionRowCell).toEqual({ kind: "checkbox" });
      expect(ledger.due?.actionRowCell).toEqual({
        kind: "groupHeading2",
        text: "Dates",
      });
    });

    it("records a spanning level-1 group heading on its first cell only", () => {
      seedFixture({
        extraSheets: [
          ledgerSheet([
            { header: "Amount", values: [10], groupHeading1: "Money" },
            { header: "Fee", values: [1] },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.groupHeading1).toBe("Money");
      expect(ledger.fee?.groupHeading1).toBeNull();
    });
  });
});
