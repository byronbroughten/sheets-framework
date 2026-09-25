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
  floorTabSeedByGid,
} from "../../01_SpreadsheetSchema/configSheetFloorSeed";
import { getSheetTraitByName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/SheetBaseNamed";
import type { SheetNamed } from "../../04_SpreadsheetNamed/SheetNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
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

interface SelfDescribingRowRule<SN extends FloorSheetName> {
  declaredColumn: ColumnName<SN>;
  identityColumns: readonly ColumnName<SN>[];
  isFloorIdentity: (identityValues: readonly CellValue[]) => boolean;
}

interface FloorTabRules<SN extends FloorSheetName> {
  excludedDataColumns: readonly ColumnName<SN>[];
  actionRowEditableColumns: readonly ColumnName<SN>[];
  selfDescribingRow: SelfDescribingRowRule<SN> | undefined;
}

export class FloorTabEditWarning<
  SN extends FloorSheetName,
> extends SheetBaseNamed<SN> {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get sheet(): SheetNamed<SN> {
    return this.ss.sheet(this.sheetName);
  }
  // Before the floor's fetch, so these ride it: a drifted column ID leaves only the Table header to find them by.
  gatherIdentityColumns(): number[] | undefined {
    const rule = floorTabRules()[this.sheetName].selfDescribingRow;
    if (rule === undefined) return undefined;
    const sheetGid = getSheetTraitByName(this.sheetName, "sheetGid");
    if (!this.ss.raw.gidIsActive(sheetGid)) return undefined;
    const sheet = this.sheet;
    if (sheet.raw.tables.length !== 1) return undefined;
    const table = sheet.raw.activeTable;
    const colIndexes = rule.identityColumns.flatMap((columnName) => {
      const header = getColumnTraitByName(this.sheetName, columnName, "header");
      const column = table.columnProperties.find(
        (colProps) => colProps.columnName === header,
      );
      return column === undefined
        ? []
        : [table.startColumnIndex + column.columnIndex];
    });
    if (colIndexes.length !== rule.identityColumns.length) return undefined;
    colIndexes.forEach((colIndex) => {
      sheet.raw.column(colIndex).gatherFetchFull();
    });
    return colIndexes;
  }
  declaration(identityColIndexes: number[] | undefined): FloorDeclaration {
    return {
      description: floorWarningDescription(this.sheetName),
      range: this.sheet.raw.wholeSheetGridRange,
      unprotectedRanges: this._editableRanges(
        this._carvedRowIndexesByColIndex(identityColIndexes),
      ),
    };
  }
  addedColumnReportLines(): string[] {
    const sheet = this.sheet;
    return this._addedColIndexes().map(
      (colIndex) =>
        `${sheet.raw.title} · ${String(sheet.raw.meta.tableHeaderRow.valueOrEmpty(colIndex))}`,
    );
  }
  queueAdd({ description, unprotectedRanges }: FloorDeclaration): void {
    this.sheet.addEditWarningWholeSheet({ description, unprotectedRanges });
  }
  // Row indexes as fetched, before the sync moves rows.
  private _carvedRowIndexesByColIndex(
    identityColIndexes: number[] | undefined,
  ): Map<number, number[]> {
    const rule = floorTabRules()[this.sheetName].selfDescribingRow;
    if (rule === undefined || identityColIndexes === undefined) {
      return new Map();
    }
    const colIndex = this._liveColIndexes().get(rule.declaredColumn);
    if (colIndex === undefined) return new Map();
    const sheet = this.sheet;
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
    const rules = floorTabRules()[this.sheetName];
    const liveIndexes = this._liveColIndexes();
    const editableDataColumns = getSheetColumnNames(this.sheetName).filter(
      (columnName) => !rules.excludedDataColumns.includes(columnName),
    );
    const sheet = this.sheet;
    const sheetId = sheet.schema.sheetGid;
    const ranges = [
      ...columnEditableRanges({
        sheetId,
        startRowIndex: sheet.schema.actionRowIndex,
        endRowIndex: sheet.schema.actionRowIndex + 1,
        colIndexes: liveColIndexesOf(
          liveIndexes,
          rules.actionRowEditableColumns,
        ),
      }),
      ...columnEditableRanges({
        sheetId,
        startRowIndex: sheet.schema.topDataRowIdx,
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
    return this.sheet.raw.fullTableColIndexes.filter(
      (colIndex) => !namedIndexes.has(colIndex),
    );
  }
  private _liveColIndexes(): Map<ColumnName<SN>, number> {
    const indexes = new Map<ColumnName<SN>, number>();
    const meta = this.sheet.raw.meta;
    getSheetColumnNames(this.sheetName).forEach((columnName) => {
      const colIndex = liveColIndex(meta, {
        columnId: getColumnTraitByName(this.sheetName, columnName, "columnId"),
        header: getColumnTraitByName(this.sheetName, columnName, "header"),
      });
      if (colIndex !== undefined) indexes.set(columnName, colIndex);
    });
    return indexes;
  }
}

export function selfDescribingRowColumns<SN extends FloorSheetName>(
  sheetName: SN,
): readonly ColumnName<SN>[] {
  const rule = floorTabRules()[sheetName].selfDescribingRow;
  if (rule === undefined) return [];
  return [...rule.identityColumns, rule.declaredColumn];
}

function floorTabRules(): { [SN in FloorSheetName]: FloorTabRules<SN> } {
  return {
    spreadsheetConfig: {
      excludedDataColumns: [
        "tableMenuSpace",
        ...spreadsheetConfigFeedbackColumnNames(),
      ],
      actionRowEditableColumns: spreadsheetConfigTimeLastRanColumnNames(),
      selfDescribingRow: undefined,
    },
    sheetConfig: {
      excludedDataColumns: ["sheetGid", "sheetTitle"],
      actionRowEditableColumns: [],
      selfDescribingRow: {
        declaredColumn: "letApiAccess",
        identityColumns: ["sheetGid"],
        isFloorIdentity: ([sheetGid]) =>
          typeof sheetGid === "number" &&
          floorTabSeedByGid(sheetGid) !== undefined,
      },
    },
    columnConfig: {
      excludedDataColumns: ["sheetGid", "columnId", "sheetTitle", "header"],
      actionRowEditableColumns: [],
      selfDescribingRow: {
        declaredColumn: "emptyValueAllowed",
        identityColumns: ["sheetGid", "columnId"],
        isFloorIdentity: ([sheetGid, columnId]) =>
          typeof sheetGid === "number" &&
          floorSeedColumnById(sheetGid, String(columnId)) !== undefined,
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

function floorWarningDescription(sheetName: FloorSheetName): string {
  return `${floorWarningPrefix} · ${configSheetFloorSeed[sheetName].title} · warning`;
}

function liveColIndexesOf<SN extends FloorSheetName>(
  liveIndexes: ReadonlyMap<ColumnName<SN>, number>,
  columnNames: readonly ColumnName<SN>[],
): number[] {
  return columnNames.flatMap((columnName) => {
    const colIndex = liveIndexes.get(columnName);
    return colIndex === undefined ? [] : [colIndex];
  });
}

interface ColumnEditableRangeProps {
  sheetId: number;
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
        startRowIndex: span.startRowIndex,
        ...(span.endRowIndex === undefined
          ? {}
          : { endRowIndex: span.endRowIndex }),
        startColumnIndex: range.startIndex,
        endColumnIndex: range.endIndex,
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
