import { extendedValue } from "../extendedValue";
import type {
  CellAfter,
  CellChange,
  DropdownStyleReset,
  GoogleGridRange,
  GridCell,
  HeldItem,
  RequestVerb,
} from "./flushClassifierTypes";
import type { BoundedRange, FlushTracker } from "./FlushTracker";
import type { RunReads } from "./RunReads";

type GoogleTable = GoogleAppsScript.Sheets.Schema.Table;

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
  readonly reads: RunReads;
  readonly tracker: FlushTracker;
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
    return boundedCells(bounded);
  }
  // What a cell holds partway through the flush, as a value it could pass on.
  afterFrom(cell: GridCell): CellAfter {
    const origin = this.tracker.resolve(cell);
    if (origin.kind === "read") {
      const read = this.reads.valueAt(origin.cell);
      if (read.kind === "read") {
        return read.value === undefined
          ? { kind: "cleared" }
          : { kind: "value", value: read.value };
      }
    }
    return { kind: "unknown", label: "(not read)" };
  }
  write(writes: CellWrite[]): this {
    writes.forEach(({ cell, after }) => this._judge(cell, after));
    this.tracker.write(writes.map(({ cell }) => cell));
    return this;
  }
  delete(cells: GridCell[]): this {
    cells.forEach((cell) => this._judge(cell, { kind: "deleted" }));
    return this;
  }
  judgeDropdownReset(table: GoogleTable): this {
    if (this.dropdownStyleReset === "allow") return this;
    const tableId = table.tableId ?? "";
    if (this.tracker.isTableAdded(tableId)) return this;
    const readTable = this.reads.table(tableId);
    if (
      readTable !== undefined &&
      !hasDropdown(readTable) &&
      !hasDropdown(table)
    ) {
      return this;
    }
    this._hold({
      kind: "dropdownReset",
      requestIndex: this.requestIndex,
      table: readTable?.name ?? table.name ?? tableId,
    });
    return this;
  }
  holdFindReplace(find: string, replacement: string): this {
    this._hold({
      kind: "findReplace",
      requestIndex: this.requestIndex,
      find,
      replacement,
    });
    return this;
  }
  holdUnmeasured(detail: string): this {
    this._hold({
      kind: "unmeasured",
      requestIndex: this.requestIndex,
      verb: this.verb,
      detail,
    });
    return this;
  }
  private _judge(cell: GridCell, after: CellAfter): void {
    const origin = this.tracker.resolve(cell);
    if (origin.kind === "created" || origin.kind === "written") return;
    if (origin.kind === "reordered") {
      this.holdUnmeasured(
        `row ${cell.rowIndex + 1}, column ${cell.colIndex + 1} on gid ${cell.sheetGid}, which a sort earlier in this flush reordered`,
      );
      return;
    }
    const read = this.reads.cellAt(origin.cell);
    if (read.kind === "outsideTables") return;
    const heldCell = this.reads.heldCell(origin.cell);
    if (read.kind === "unread") {
      this._hold({
        kind: "unreadCell",
        requestIndex: this.requestIndex,
        verb: this.verb,
        cell: heldCell,
      });
      return;
    }
    const before = read.value;
    if (before === undefined || before.formulaValue !== undefined) return;
    if (after.kind === "value" && extendedValue.isEqual(after.value, before)) {
      return;
    }
    this._hold({
      kind: "enteredData",
      requestIndex: this.requestIndex,
      verb: this.verb,
      cell: heldCell,
      change: cellChange(extendedValue.label(before), after),
    });
  }
  private _hold(item: HeldItem): void {
    this.held.push(item);
  }
}

export function boundedCells(range: BoundedRange): GridCell[] {
  const { sheetGid, rows, cols } = range;
  return spanIndexes(rows).flatMap((rowIndex) =>
    spanIndexes(cols).map((colIndex) => ({ sheetGid, rowIndex, colIndex })),
  );
}

function spanIndexes(span: { startIndex: number; endIndex: number }): number[] {
  return Array.from(
    { length: Math.max(span.endIndex - span.startIndex, 0) },
    (_, offset) => span.startIndex + offset,
  );
}

function hasDropdown(table: GoogleTable): boolean {
  return (table.columnProperties ?? []).some(
    (column) =>
      column.columnType === "DROPDOWN" ||
      column.dataValidationRule !== undefined,
  );
}

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
        after: extendedValue.label(after.value),
      };
    case "unknown":
      return { kind: "overwrite", before, after: after.label };
    default: {
      const exhaustive: never = after;
      throw new Error(`Unknown cell after: ${JSON.stringify(exhaustive)}`);
    }
  }
}
