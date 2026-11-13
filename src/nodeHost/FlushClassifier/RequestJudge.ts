import { ExtVal } from "../ExtVal";
import type {
  BoundedRange,
  CellAfter,
  CellChange,
  DropdownStyleReset,
  GoogleGridRange,
  GridCell,
  HeldItem,
  RequestVerb,
} from "./flushClassifierTypes";
import type { FlushTracker } from "./FlushTracker";
import { Grid } from "./Grid";
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
    cells.forEach((cell) => this._judge(cell, { kind: "deleted" }));
    return this;
  }
  // Every cell of the read range, header included, since deleteTable blanks it all.
  deleteTable(tableId: string): this {
    if (this.tracker.isTableAdded(tableId)) return this;
    const range = this.reads.table(tableId)?.range;
    if (range === undefined) {
      return this.holdUnmeasured(`Table ${tableId}, which no read fetched`);
    }
    const cells = this.cellsIn(range);
    this.delete(cells);
    this.tracker.write(cells);
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
    this.held.push({
      kind: "dropdownReset",
      requestIndex: this.requestIndex,
      table: readTable?.name ?? table.name ?? tableId,
    });
    return this;
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
      this.held.push({
        kind: "unreadCell",
        requestIndex: this.requestIndex,
        verb: this.verb,
        cell: heldCell,
      });
      return;
    }
    const before = read.value;
    if (before === undefined || before.formulaValue !== undefined) return;
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

const notReadAfter: CellAfter = { kind: "unknown", label: "(not read)" };

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
