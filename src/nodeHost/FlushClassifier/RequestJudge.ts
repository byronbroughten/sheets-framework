import type { DropdownStyleReset } from "../../00_Source/RawSource/RawSource.ts";
import { ExtVal } from "../ExtVal.ts";
import type {
  BoundedRange,
  CellAfter,
  CellChange,
  GoogleGridRange,
  GridCell,
  HeldItem,
  RequestVerb,
} from "./flushClassifierTypes.ts";
import type { FlushTracker } from "./FlushTracker.ts";
import { Grid } from "./Grid.ts";
import type { ReadValue, RunReads } from "./RunReads.ts";

type GoogleTable = GoogleAppsScript.Sheets.Schema.Table;
type GoogleTableColumnProperties =
  GoogleAppsScript.Sheets.Schema.TableColumnProperties;

export interface CellWrite {
  cell: GridCell;
  after: CellAfter;
}

export interface RequestJudgeProps {
  reads: RunReads;
  tracker: FlushTracker;
  requestIndex: number;
  verb: RequestVerb;
  dropdownStyleReset: DropdownStyleReset;
}

// Judges one request's cells against the run's reads and gathers what it holds.
export class RequestJudge {
  private reads: RunReads;
  private tracker: FlushTracker;
  private requestIndex: number;
  private verb: RequestVerb;
  private dropdownStyleReset: DropdownStyleReset;
  private held: HeldItem[];
  constructor(props: RequestJudgeProps) {
    this.reads = props.reads;
    this.tracker = props.tracker;
    this.requestIndex = props.requestIndex;
    this.verb = props.verb;
    this.dropdownStyleReset = props.dropdownStyleReset;
    this.held = [];
  }
  get heldItems(): HeldItem[] {
    return [...this.held];
  }
  // Undefined, and held as unmeasured, when an open bound meets a tab no read sized.
  bounded(range: GoogleGridRange): BoundedRange | undefined {
    const bounded = this.tracker.bounded(range);
    if (bounded === undefined) {
      this.holdUnmeasured(
        `an open-ended range on gid ${range.sheetId ?? 0}, a tab no read sized`,
      );
    }
    return bounded;
  }
  cellsIn(range: GoogleGridRange): GridCell[] {
    const bounded = this.bounded(range);
    if (bounded === undefined) return [];
    return Grid.cells(bounded);
  }
  // What a cell holds partway through the flush, as a value it could pass on.
  afterFrom(cell: GridCell): CellAfter {
    const origin = this.tracker.resolve(cell);
    if (origin.kind !== "read") return notReadAfter;
    const read = this.reads.valueAt(origin.cell);
    if (read.kind === "unread") return notReadAfter;
    if (read.value === undefined) return { kind: "cleared" };
    return { kind: "value", value: read.value };
  }
  write(writes: CellWrite[]): this {
    writes.forEach(({ cell, after }) => this._judge(cell, after));
    this.tracker.write(writes.map(({ cell }) => cell));
    return this;
  }
  delete(cells: GridCell[]): this {
    for (const cell of cells) {
      this._judge(cell, deletedAfter);
    }
    return this;
  }
  // Every held cell on the tab, which only a read of its whole used grid can show.
  deleteSheet(sheetGid: number): this {
    if (!this.reads.hasUsedGrid(sheetGid)) {
      return this.holdUnmeasured(
        `tab gid ${sheetGid}, deleted with every Table on it`,
      );
    }
    for (const cell of this.cellsIn({ sheetId: sheetGid })) {
      this._judgeTabCell(cell);
    }
    return this;
  }
  // Every cell of the read range where earlier requests left it, header included, since deleteTable blanks it all.
  deleteTable(tableId: string): this {
    if (this.tracker.isTableAdded(tableId)) return this;
    const range = this.reads.table(tableId)?.range;
    if (range === undefined) {
      return this.holdUnmeasured(`Table ${tableId}, which no read fetched`);
    }
    const tableNow = this._tableNow(range);
    if (tableNow.kind === "unfollowed") {
      return this.holdUnmeasured(`Table ${tableId}, ${tableNow.reason}`);
    }
    const cells = Grid.cells(tableNow.now);
    this.delete(cells);
    this.tracker.write(cells);
    return this;
  }
  judgeDropdownReset(table: GoogleTable): this {
    if (this.dropdownStyleReset === "allow") return this;
    const tableId = table.tableId ?? "";
    if (this.tracker.isTableAdded(tableId)) return this;
    const readTable = this.reads.table(tableId);
    const readDropdowns = this._readDropdownsNow(readTable);
    if (
      readDropdowns !== undefined &&
      !dropdowns.isLeftOut(readDropdowns, table) &&
      !dropdowns.isIn(table)
    ) {
      return this;
    }
    this.held.push({
      kind: "dropdownReset",
      requestIndex: this.requestIndex,
      table: readTable?.name ?? table.name ?? tableId,
    });
    return this;
  }
  // Each read dropdown's column index within its Table where earlier requests left it; undefined when the tracker can't follow the Table or one of its dropdowns.
  private _readDropdownsNow(
    read: GoogleTable | undefined,
  ): number[] | undefined {
    if (read?.range === undefined) return undefined;
    const tableNow = this._tableNow(read.range);
    if (tableNow.kind === "unfollowed") return undefined;
    const indexes: number[] = [];
    const readDropdowns = (read.columnProperties ?? []).filter(
      dropdowns.isColumn,
    );
    for (const { columnIndex = 0 } of readDropdowns) {
      const colIndex = tableNow.read.cols.startIndex + columnIndex;
      const colNow = this.tracker.follow({
        ...tableNow.read,
        cols: { startIndex: colIndex, endIndex: colIndex + 1 },
      });
      if (colNow === undefined || Grid.isSpanEmpty(colNow.cols)) {
        return undefined;
      }
      indexes.push(colNow.cols.startIndex - tableNow.now.cols.startIndex);
    }
    return indexes;
  }
  // A read Table's range, bounded by its tab's size as read, and where earlier requests left it.
  private _tableNow(range: GoogleGridRange): TableNow {
    const sheetGid = range.sheetId ?? 0;
    const read = Grid.bounded(range, this.reads.gridSizes.get(sheetGid));
    if (read === undefined) {
      return {
        kind: "unfollowed",
        reason: `whose read range is open-ended on gid ${sheetGid}, a tab no read sized`,
      };
    }
    const now = this.tracker.follow(read);
    if (now === undefined) {
      return {
        kind: "unfollowed",
        reason: "whose read range an earlier request split",
      };
    }
    return { kind: "followed", read, now };
  }
  holdFindReplace(find: string, replacement: string): this {
    this.held.push({
      kind: "findReplace",
      requestIndex: this.requestIndex,
      find,
      replacement,
    });
    return this;
  }
  holdUnmeasured(detail: string): this {
    this.held.push({
      kind: "unmeasured",
      requestIndex: this.requestIndex,
      verb: this.verb,
      detail,
    });
    return this;
  }
  private _judge(cell: GridCell, after: CellAfter): void {
    const origin = this._readOrigin(cell);
    if (origin === undefined) return;
    const read = this.reads.cellAt(origin);
    if (read.kind === "outsideTables") return;
    this._holdEnteredData({ origin, read, after });
  }
  // A Table's body cell is judged as deleteTable judges it, even inside another Table's head rows; any other non-blank cell, a formula included, is held.
  private _judgeTabCell(cell: GridCell): void {
    const origin = this._readOrigin(cell);
    if (origin === undefined) return;
    const read = this.reads.cellAt(origin);
    if (read.kind !== "outsideTables") {
      this._holdEnteredData({ origin, read, after: deletedAfter });
      return;
    }
    if (this.reads.isInHeadRows(origin)) return;
    const value = this.reads.valueAt(origin);
    this._holdChange({ origin, read: value, after: deletedAfter });
  }
  // The read cell a cell holds; undefined when the flush made it, or a sort hid which.
  private _readOrigin(cell: GridCell): GridCell | undefined {
    const origin = this.tracker.resolve(cell);
    if (origin.kind === "created" || origin.kind === "written") {
      return undefined;
    }
    if (origin.kind === "reordered") {
      this.holdUnmeasured(
        `row ${cell.rowIndex + 1}, column ${cell.colIndex + 1} on gid ${cell.sheetGid}, which a sort earlier in this flush reordered`,
      );
      return undefined;
    }
    return origin.cell;
  }
  // A formula in a Table's body isn't entered data.
  private _holdEnteredData(judged: JudgedCell): void {
    const { read } = judged;
    if (read.kind === "read" && read.value?.formulaValue !== undefined) return;
    this._holdChange(judged);
  }
  private _holdChange({ origin, read, after }: JudgedCell): void {
    const heldCell = this.reads.heldCell(origin);
    if (read.kind === "unread") {
      this.held.push({
        kind: "unreadCell",
        requestIndex: this.requestIndex,
        verb: this.verb,
        cell: heldCell,
      });
      return;
    }
    const before = read.value;
    if (before === undefined) return;
    if (after.kind === "value" && ExtVal.isEqual(after.value, before)) {
      return;
    }
    this.held.push({
      kind: "enteredData",
      requestIndex: this.requestIndex,
      verb: this.verb,
      cell: heldCell,
      change: cellChange(ExtVal.label(before), after),
    });
  }
}

type TableNow =
  | { kind: "followed"; read: BoundedRange; now: BoundedRange }
  | { kind: "unfollowed"; reason: string };

interface JudgedCell {
  origin: GridCell;
  read: ReadValue;
  after: CellAfter;
}

const notReadAfter: CellAfter = { kind: "unknown", label: "(not read)" };
const deletedAfter: CellAfter = { kind: "deleted" };

const dropdowns = {
  isIn(table: GoogleTable): boolean {
    return (table.columnProperties ?? []).some(dropdowns.isColumn);
  },
  // A read dropdown the sent Table retypes is being removed, with no style to keep; one it leaves out is held, since this judge doesn't model what a partial list does to it.
  isLeftOut(readDropdowns: number[], sent: GoogleTable): boolean {
    const sentColumns = sent.columnProperties ?? [];
    return readDropdowns.some(
      (readIndex) =>
        !sentColumns.some(({ columnIndex }) => columnIndex === readIndex),
    );
  },
  isColumn(column: GoogleTableColumnProperties): boolean {
    return (
      column.columnType === "DROPDOWN" ||
      column.dataValidationRule !== undefined
    );
  },
};

const emptyLabel = "(empty)";

function cellChange(before: string, after: CellAfter): CellChange {
  switch (after.kind) {
    case "deleted":
      return { kind: "delete", before };
    case "cleared":
      return { kind: "overwrite", before, after: emptyLabel };
    case "value":
      return {
        kind: "overwrite",
        before,
        after: ExtVal.label(after.value),
      };
    case "unknown":
      return { kind: "overwrite", before, after: after.label };
    default: {
      const exhaustive: never = after;
      throw new Error(`Unknown cell after: ${JSON.stringify(exhaustive)}`);
    }
  }
}
