import type { CellValue } from "../../00_Source/CellValues/cellValues";
import type { ProtectionGridRange } from "../../00_Source/RawSource/EditProtection";
import {
  type ColumnName,
  getColumnTraitByName,
  getSheetColumnNames,
} from "../../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  floorSeedColumnById,
  floorTabSeedByTableId,
} from "../../01_SpreadsheetSchema/configSheetFloorSeed";
import { getTableTraitByName } from "../../01_SpreadsheetSchema/tableConfigsTypes";
import type { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { TableBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/TableBaseNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import type { TableNamed } from "../../04_SpreadsheetNamed/TableNamed";
import { Arr } from "../../utils/Arr";
import { Obj } from "../../utils/Obj";
import { liveColIndex } from "./floorColumnLocation";
import {
  columnNameByHeader,
  type FloorSheetName,
  spreadsheetConfigFeedbackColumnNames,
} from "./floorSeedLookups";

export const floorWarningPrefix = "Config-sheet floor";

export interface FloorDeclaration {
  description: string;
  range: ProtectionGridRange;
  unprotectedRanges: ProtectionGridRange[];
}

interface SelfDescribingRowRule<TN extends FloorSheetName> {
  declaredColumn: ColumnName<TN>;
  identityColumns: readonly ColumnName<TN>[];
  isFloorIdentity: (identityValues: readonly CellValue[]) => boolean;
}

interface FloorTabRules<TN extends FloorSheetName> {
  excludedDataColumns: readonly ColumnName<TN>[];
  actionRowEditableColumns: readonly ColumnName<TN>[];
  selfDescribingRow: SelfDescribingRowRule<TN> | undefined;
}

export class FloorTabEditWarning<
  TN extends FloorSheetName,
> extends TableBaseNamed<TN> {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get table(): TableNamed<TN> {
    return this.ss.table(this.tableName);
  }
  // Before the floor's fetch, so these ride it: a drifted column ID leaves only the Table header to find them by.
  gatherIdentityColumns(): number[] | undefined {
    const rule = floorTabRules()[this.tableName].selfDescribingRow;
    if (rule === undefined) return undefined;
    const sheetGid = getTableTraitByName(this.tableName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return undefined;
    const sheet = this.table;
    if (!sheet.raw.hasOneTable()) return undefined;
    const table = sheet.raw;
    const colIndexes = rule.identityColumns.flatMap((columnName) => {
      const header = getColumnTraitByName(this.tableName, columnName, "header");
      const column = table.columnProperties.find(
        (colProps) => colProps.columnName === header,
      );
      return column === undefined ? [] : [column.columnIndex];
    });
    if (colIndexes.length !== rule.identityColumns.length) return undefined;
    colIndexes.forEach((colIndex) => {
      sheet.raw.column(colIndex).gatherFetchFull();
    });
    return colIndexes;
  }
  declaration(identityColIndexes: number[] | undefined): FloorDeclaration {
    return {
      description: floorWarningDescription(this.tableName),
      range: this.table.raw.sheet.wholeSheetGridRange,
      unprotectedRanges: this._editableRanges(
        this._carvedRowIndexesByColIndex(identityColIndexes),
      ),
    };
  }
  addedColumnReportLines(): string[] {
    const sheet = this.table;
    return this._addedColIndexes().map(
      (colIndex) =>
        `${sheet.raw.sheet.title} · ${String(sheet.raw.headRow("header").valueOrEmpty(colIndex))}`,
    );
  }
  queueAdd({ description, unprotectedRanges }: FloorDeclaration): void {
    this.table.sheet.addEditWarningWholeSheet({
      description,
      unprotectedRanges,
    });
  }
  // Row indexes as fetched, before the sync moves rows.
  private _carvedRowIndexesByColIndex(
    identityColIndexes: number[] | undefined,
  ): Map<number, number[]> {
    const rule = floorTabRules()[this.tableName].selfDescribingRow;
    if (rule === undefined || identityColIndexes === undefined) {
      return new Map();
    }
    const colIndex = this._liveColIndexes().get(rule.declaredColumn);
    if (colIndex === undefined) return new Map();
    const sheet = this.table;
    return new Map([
      [
        colIndex,
        sheet.raw.rowIndexesFull.filter((rowIndex) =>
          rule.isFloorIdentity(
            identityColIndexes.map((identityColIndex) =>
              sheet.raw.column(identityColIndex).valueOrEmpty(rowIndex),
            ),
          ),
        ),
      ],
    ]);
  }
  private _editableRanges(
    carvedRowIndexesByColIndex: ReadonlyMap<number, readonly number[]>,
  ): ProtectionGridRange[] {
    const rules = floorTabRules()[this.tableName];
    const liveIndexes = this._liveColIndexes();
    const editableDataColumns = getSheetColumnNames(this.tableName).filter(
      (columnName) => !rules.excludedDataColumns.includes(columnName),
    );
    const sheet = this.table;
    const sheetId = sheet.schema.sheetGid;
    const origin = sheet.raw.tableOrigin();
    const ranges = [
      ...columnEditableRanges({
        sheetId,
        origin,
        startRowIndex: sheet.schema.actionRowIndex,
        endRowIndex: sheet.schema.actionRowIndex + 1,
        colIndexes: liveColIndexesOf(
          liveIndexes,
          rules.actionRowEditableColumns,
        ),
      }),
      ...columnEditableRanges({
        sheetId,
        origin,
        startRowIndex: 0,
        colIndexes: [
          ...liveColIndexesOf(liveIndexes, editableDataColumns),
          ...this._addedColIndexes(),
        ],
        carvedRowIndexesByColIndex,
      }),
    ];
    return ranges.sort(compareProtectionRanges);
  }
  private _addedColIndexes(): number[] {
    const namedIndexes = new Set(this._liveColIndexes().values());
    return this.table.raw.fullTableColIndexes.filter(
      (colIndex) => !namedIndexes.has(colIndex),
    );
  }
  private _liveColIndexes(): Map<ColumnName<TN>, number> {
    const indexes = new Map<ColumnName<TN>, number>();
    const rawTable = this.table.raw;
    getSheetColumnNames(this.tableName).forEach((columnName) => {
      const colIndex = liveColIndex(rawTable, {
        columnId: getColumnTraitByName(this.tableName, columnName, "columnId"),
        header: getColumnTraitByName(this.tableName, columnName, "header"),
      });
      if (colIndex !== undefined) indexes.set(columnName, colIndex);
    });
    return indexes;
  }
}

export function selfDescribingRowColumns<TN extends FloorSheetName>(
  tableName: TN,
): readonly ColumnName<TN>[] {
  const rule = floorTabRules()[tableName].selfDescribingRow;
  if (rule === undefined) return [];
  return [...rule.identityColumns, rule.declaredColumn];
}

function floorTabRules(): { [TN in FloorSheetName]: FloorTabRules<TN> } {
  return {
    spreadsheetConfig: {
      excludedDataColumns: [
        "tableMenuSpace",
        ...spreadsheetConfigFeedbackColumnNames(),
      ],
      actionRowEditableColumns: spreadsheetConfigTimeLastRanColumnNames(),
      selfDescribingRow: undefined,
    },
    tableConfig: {
      excludedDataColumns: ["tableId", "tableName", "sheetTitle"],
      actionRowEditableColumns: [],
      selfDescribingRow: {
        declaredColumn: "letApiAccess",
        identityColumns: ["tableId"],
        isFloorIdentity: ([tableId]) =>
          typeof tableId === "string" &&
          floorTabSeedByTableId(tableId) !== undefined,
      },
    },
    columnConfig: {
      excludedDataColumns: ["tableId", "columnId", "tableName", "header"],
      actionRowEditableColumns: [],
      selfDescribingRow: {
        declaredColumn: "emptyValueAllowed",
        identityColumns: ["tableId", "columnId"],
        isFloorIdentity: ([tableId, columnId]) =>
          typeof tableId === "string" &&
          floorSeedColumnById(tableId, String(columnId)) !== undefined,
      },
    },
  };
}

function spreadsheetConfigTimeLastRanColumnNames(): ColumnName<"spreadsheetConfig">[] {
  return Obj.values(configSheetFloorSeed.spreadsheetConfig.endpoints).map(
    (endpoint) =>
      columnNameByHeader("spreadsheetConfig", endpoint.timeLastRan.header),
  );
}

function floorWarningDescription(tableName: FloorSheetName): string {
  return `${floorWarningPrefix} · ${configSheetFloorSeed[tableName].title} · warning`;
}

function liveColIndexesOf<TN extends FloorSheetName>(
  liveIndexes: ReadonlyMap<ColumnName<TN>, number>,
  columnNames: readonly ColumnName<TN>[],
): number[] {
  return columnNames.flatMap((columnName) => {
    const colIndex = liveIndexes.get(columnName);
    return colIndex === undefined ? [] : [colIndex];
  });
}

// Table-relative rows and columns, converted to the sheet's by origin.
interface ColumnEditableRangeProps {
  sheetId: number;
  origin: TableOrigin;
  startRowIndex: number;
  endRowIndex?: number;
  colIndexes: number[];
  carvedRowIndexesByColIndex?: ReadonlyMap<number, readonly number[]>;
}

interface RowSpan {
  startRowIndex: number;
  endRowIndex?: number;
}

function columnEditableRanges({
  sheetId,
  origin,
  startRowIndex,
  endRowIndex,
  colIndexes,
  carvedRowIndexesByColIndex = new Map(),
}: ColumnEditableRangeProps): ProtectionGridRange[] {
  const colIndexesBySpans = new Map<
    string,
    { spans: RowSpan[]; colIndexes: number[] }
  >();
  colIndexes.forEach((colIndex) => {
    const spans = uncarvedRowSpans({
      startRowIndex,
      endRowIndex,
      carvedRowIndexes: carvedRowIndexesByColIndex.get(colIndex) ?? [],
    });
    const key = spansKey(spans);
    const group = colIndexesBySpans.get(key) ?? { spans, colIndexes: [] };
    group.colIndexes.push(colIndex);
    colIndexesBySpans.set(key, group);
  });
  return [...colIndexesBySpans.values()].flatMap((group) =>
    Arr.contiguousRanges(group.colIndexes).flatMap((range) =>
      group.spans.map((span) => ({
        sheetId,
        startRowIndex: origin.sheetRowIndex(span.startRowIndex),
        ...(span.endRowIndex === undefined
          ? {}
          : { endRowIndex: origin.sheetRowIndex(span.endRowIndex) }),
        startColumnIndex: origin.sheetColIndex(range.startIndex),
        endColumnIndex: origin.sheetColIndex(range.endIndex),
      })),
    ),
  );
}

function uncarvedRowSpans({
  startRowIndex,
  endRowIndex,
  carvedRowIndexes,
}: {
  startRowIndex: number;
  endRowIndex?: number;
  carvedRowIndexes: readonly number[];
}): RowSpan[] {
  const carved = Arr.contiguousRanges(
    carvedRowIndexes.filter(
      (rowIndex) =>
        rowIndex >= startRowIndex &&
        (endRowIndex === undefined || rowIndex < endRowIndex),
    ),
  );
  const spanStarts = [startRowIndex, ...carved.map((range) => range.endIndex)];
  const spanEnds = [...carved.map((range) => range.startIndex), endRowIndex];
  return spanStarts.flatMap((start, i) => {
    const end = spanEnds[i];
    if (end !== undefined && end <= start) return [];
    return [
      end === undefined
        ? { startRowIndex: start }
        : { startRowIndex: start, endRowIndex: end },
    ];
  });
}

// Columns with identical spans share a key, so their ranges merge.
function spansKey(spans: readonly RowSpan[]): string {
  return JSON.stringify(spans);
}

function compareProtectionRanges(
  left: ProtectionGridRange,
  right: ProtectionGridRange,
): number {
  const leftRow = "startRowIndex" in left ? left.startRowIndex : -1;
  const rightRow = "startRowIndex" in right ? right.startRowIndex : -1;
  if (leftRow !== rightRow) return leftRow - rightRow;
  const leftCol = "startColumnIndex" in left ? (left.startColumnIndex ?? 0) : 0;
  const rightCol =
    "startColumnIndex" in right ? (right.startColumnIndex ?? 0) : 0;
  return leftCol - rightCol;
}
