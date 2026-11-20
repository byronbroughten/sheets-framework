import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCellValue,
  type FakeSheetProperties,
} from "../testSupport/fakeSheetsService";
import {
  type ConfigRegeneration,
  SyncCoordinator,
  type SyncRun,
} from "./SyncCoordinator";
import {
  frame,
  previousItemMap,
  seedFixture,
  sendWithNoAlterations,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator send with no alterations head-row overlap", () => {
  const recordsGid = 888000222;

  interface RecordsTable {
    name: string;
    headerRowIndex: number;
    startColumnIndex: number;
    endColumnIndex: number;
    endRowIndex: number;
    isManaged: boolean;
  }

  // A managed Table takes a Table ID the previous configs grant.
  const grantedTableIdsByName: Record<string, string> = {
    Leases: getTableTraitByName("log", "tableId"),
    Tenants: getTableTraitByName("runItem", "tableId"),
    Units: getTableTraitByName("computed", "tableId"),
    Owners: getTableTraitByName("dates", "tableId"),
  };

  function recordsTableId(table: RecordsTable): string {
    if (!table.isManaged) return `records-${table.name.toLowerCase()}`;
    return Val.assert(
      grantedTableIdsByName[table.name],
      "granted Table ID of a managed Records Table",
    );
  }

  // Several Tables on one sheet, which the one-Table builder can't place.
  function recordsSheet(tables: readonly RecordsTable[]): FakeSheetProperties {
    const rows: Record<number, FakeCellValue[]> = {};
    tables.forEach((table) => {
      const headerRow = (rows[table.headerRowIndex] ??= []);
      const topRow = (rows[table.headerRowIndex + 1] ??= []);
      for (
        let col = table.startColumnIndex;
        col < table.endColumnIndex;
        col++
      ) {
        headerRow[col] = `${table.name} ${col}`;
        topRow[col] ??= "";
      }
    });
    return {
      sheetId: recordsGid,
      title: "Records",
      rows: buildGridRows(
        Object.fromEntries(
          Object.entries(rows).map(([rowIndex, row]) => [
            rowIndex,
            Array.from(row, (cell) => cell ?? ""),
          ]),
        ),
      ),
      tables: tables.map((table) => ({
        tableId: recordsTableId(table),
        name: table.name,
        startRowIndex: table.headerRowIndex,
        startColumnIndex: table.startColumnIndex,
        endRowIndex: table.endRowIndex,
        endColumnIndex: table.endColumnIndex,
      })),
    };
  }

  function seedRecords(tables: readonly RecordsTable[]) {
    return seedFixture({ extraSheets: [recordsSheet(tables)] });
  }

  // The managed Tables' blank column ID cells are what a run that went on would write.
  function expectRefused(
    tables: readonly RecordsTable[],
    ...messages: string[]
  ): void {
    const { grid } = seedRecords(tables);
    const recordsValues = () => grid.sheet(recordsGid).values();
    const before = recordsValues();

    let parsed: ConfigRegeneration | undefined;
    expect(() => {
      parsed = sendWithNoAlterations();
    }).toThrow(messages.join(" "));
    expect(parsed).toBeUndefined();
    expect(recordsValues()).toEqual(before);
  }

  // Rows 4–9 are its body, 0-based.
  const leases: RecordsTable = {
    name: "Leases",
    headerRowIndex: 3,
    startColumnIndex: 1,
    endColumnIndex: 4,
    endRowIndex: 10,
    isManaged: true,
  };

  it("refuses a Table stacked under another and starting left of it, its head rows on the other's last body row", () => {
    expectRefused(
      [
        leases,
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 14,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("refuses a Table whose head rows sit on another managed Table's head rows", () => {
    expectRefused(
      [
        {
          name: "Leases",
          headerRowIndex: 3,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 5,
          isManaged: true,
        },
        {
          name: "Tenants",
          headerRowIndex: 5,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 7,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 3–5 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("refuses a managed Table whose head rows sit on a Table the app doesn't manage", () => {
    expectRefused(
      [
        { ...leases, isManaged: false },
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 14,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("lists two overlaps on one sheet in one message", () => {
    expectRefused(
      [
        {
          name: "Leases",
          headerRowIndex: 3,
          startColumnIndex: 0,
          endColumnIndex: 2,
          endRowIndex: 10,
          isManaged: true,
        },
        {
          name: "Units",
          headerRowIndex: 3,
          startColumnIndex: 3,
          endColumnIndex: 5,
          endRowIndex: 10,
          isManaged: true,
        },
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 2,
          endRowIndex: 14,
          isManaged: true,
        },
        {
          name: "Owners",
          headerRowIndex: 11,
          startColumnIndex: 3,
          endColumnIndex: 5,
          endRowIndex: 13,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
      `Table "Owners" on sheet "Records" has head rows in rows 9–11 that sit on Table "Units". Move "Owners" so that the 3 rows above its header sit clear of "Units".`,
    );
  });

  it("refuses a granted Table stacked clear of another but below the header zone, naming it", () => {
    seedRecords([
      leases,
      {
        name: "Tenants",
        headerRowIndex: 13,
        startColumnIndex: 0,
        endColumnIndex: 3,
        endRowIndex: 15,
        isManaged: true,
      },
    ]);

    expect(() => sendWithNoAlterations()).toThrowError(
      /^Table "Tenants" on "Records" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("refuses a recorded Table moved below the header zone with generation's own fix", () => {
    seedFixture({
      item: { startRowIndex: expectedSheetLayout.tableHeaderRowIndex + 2 },
    });

    expect(() => sendWithNoAlterations()).toThrowError(
      /^Table "Item" on "Item" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("leaves an ungranted Table below the header zone to the operator", () => {
    seedRecords([
      leases,
      {
        name: "Tenants",
        headerRowIndex: 13,
        startColumnIndex: 0,
        endColumnIndex: 3,
        endRowIndex: 15,
        isManaged: false,
      },
    ]);

    const parsed = sendWithNoAlterations();
    expect(parsed.tableConfigs).toContain('"leases"');
    expect(parsed.tableConfigs).not.toContain('"tenants"');
  });

  it("allows two Tables side by side, touching", () => {
    seedRecords([
      {
        name: "Leases",
        headerRowIndex: 3,
        startColumnIndex: 0,
        endColumnIndex: 2,
        endRowIndex: 6,
        isManaged: true,
      },
      {
        name: "Tenants",
        headerRowIndex: 3,
        startColumnIndex: 2,
        endColumnIndex: 4,
        endRowIndex: 5,
        isManaged: true,
      },
    ]);

    const parsed = sendWithNoAlterations();
    expect(parsed.tableConfigs).toContain('"tenants"');
    expect(parsed.tableConfigs).toContain('"leases"');
  });

  describe("granting API access", () => {
    // Tenants, ungranted, with its header on row 4.
    const tenants: RecordsTable = {
      name: "Tenants",
      headerRowIndex: 3,
      startColumnIndex: 0,
      endColumnIndex: 3,
      endRowIndex: 5,
      isManaged: false,
    };

    // The previous map holds Tenants with no columns: it has no API access.
    function previousMapWithTenants(): IdNameMapGeneric {
      const map = previousItemMap();
      const tableId = recordsTableId(tenants);
      map.tables[tableId] = { name: "tenants", sheetGid: recordsGid };
      map.tabs[String(recordsGid)] = {
        name: "records",
        title: "Records",
        tableIds: [tableId],
      };
      return map;
    }

    function runGrant(isSending: boolean): SyncRun {
      return SyncCoordinator.init().run({
        isSending,
        intendedState: {
          alterations: {
            tables: { tenants: { letApiAccess: true } },
          } as PendingAlterations,
        },
        previousIdNameMap: previousMapWithTenants(),
        forcedTraits: [],
        frame,
      });
    }

    // Notes sits in rows 1–2, under Tenants' head rows.
    const notesUnderHeadRows: RecordsTable = {
      name: "Notes",
      headerRowIndex: 0,
      startColumnIndex: 2,
      endColumnIndex: 4,
      endRowIndex: 2,
      isManaged: false,
    };

    const overlapStop = {
      kind: "headRows",
      table: "tenants",
      tab: "Records",
      neighbour: "Notes",
    };

    it("stops a grant whose head rows would sit on another Table on a dry run, naming both", () => {
      const { batchUpdateCount } = seedRecords([tenants, notesUnderHeadRows]);

      const { result } = runGrant(false);

      expect(batchUpdateCount()).toBe(0);
      expect(result).toMatchObject({
        outcome: { kind: "stopped" },
        stops: [overlapStop],
      });
    });

    it("stops the same grant on a send, sending nothing", () => {
      const { grid, batchUpdateCount } = seedRecords([
        tenants,
        notesUnderHeadRows,
      ]);
      const before = grid.sheet(recordsGid).values();

      const { result, files } = runGrant(true);

      expect(batchUpdateCount()).toBe(0);
      expect(grid.sheet(recordsGid).values()).toEqual(before);
      expect(files?.tableConfigs).not.toContain('"tenants"');
      expect(result).toMatchObject({
        outcome: { kind: "stopped" },
        landed: [],
        stops: [overlapStop],
      });
    });

    it("plans a grant whose head rows sit clear of every other Table", () => {
      seedRecords([
        tenants,
        {
          name: "Notes",
          headerRowIndex: 0,
          startColumnIndex: 3,
          endColumnIndex: 5,
          endRowIndex: 2,
          isManaged: false,
        },
      ]);

      const { result } = runGrant(false);

      expect(result).toMatchObject({
        outcome: { kind: "ready" },
        stops: [],
        plan: [
          {
            subject: { table: "tenants", path: "letApiAccess" },
            what: "set to true",
          },
        ],
      });
    });
  });
});
