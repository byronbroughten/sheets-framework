import {
  moveBeforePath,
  moveColumnPath,
  newColumnsBeforesOf,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncClashingPlacements, SyncRefusal } from "../syncRunResult";

interface NamedAt {
  column: string;
  path: string;
}

interface ClashProps {
  clash: SyncClashingPlacements["clash"];
  named: NamedAt;
  otherPath: string;
}

interface SameColumnClashProps {
  clash: SyncClashingPlacements["clash"];
  named: NamedAt;
  others: NamedAt[];
}

// Refuses a Table entry's `moves` groups and `newColumns` placements that can't all hold at once, naming both paths of each clash.
export class PlacementsRefuser {
  private table: string;
  private entry: WideTableEntry;
  constructor(table: string, entry: WideTableEntry) {
    this.table = table;
    this.entry = entry;
  }
  get refusals(): SyncRefusal[] {
    return [
      ...this._sameBeforeRefusals,
      ...this._repeatedColumnRefusals,
      ...this._movedBeforeRefusals,
      ...this._repeatRefusals(this._newColumnsBefores, "sameNewColumnsBefore"),
      ...this._newColumnsBeforeRefusals,
    ];
  }
  // New columns go in before the moves, so a move to or of their `before` leaves them no longer right before it.
  private get _newColumnsBeforeRefusals(): SyncClashingPlacements[] {
    const moveBefores = this._befores;
    const movedColumns = this._movedColumns;
    return this._newColumnsBefores.flatMap((named) => [
      ...this._sameColumnClashes({
        clash: "newColumnsBeforeMoveTarget",
        named,
        others: moveBefores,
      }),
      ...this._sameColumnClashes({
        clash: "newColumnsBeforeMovedColumn",
        named,
        others: movedColumns,
      }),
    ]);
  }
  private get _sameBeforeRefusals(): SyncClashingPlacements[] {
    return this._repeatRefusals(this._befores, "sameBefore");
  }
  private get _repeatedColumnRefusals(): SyncClashingPlacements[] {
    return this._repeatRefusals(this._movedColumns, "repeatedColumn");
  }
  // A group's target can't move, its own group's included, or one group's landing would undo another's.
  private get _movedBeforeRefusals(): SyncClashingPlacements[] {
    const movedColumns = this._movedColumns;
    return this._befores.flatMap((named) =>
      this._sameColumnClashes({
        clash: "movedBefore",
        named,
        others: movedColumns,
      }),
    );
  }
  private _sameColumnClashes({
    clash,
    named,
    others,
  }: SameColumnClashProps): SyncClashingPlacements[] {
    const other = others.find(({ column }) => column === named.column);
    if (other === undefined) return [];
    return [this._clash({ clash, named, otherPath: other.path })];
  }
  private _repeatRefusals(
    named: NamedAt[],
    clash: SyncClashingPlacements["clash"],
  ): SyncClashingPlacements[] {
    return named.flatMap((repeat, index) => {
      const firstIndex = named.findIndex(
        ({ column }) => column === repeat.column,
      );
      const first = named[firstIndex];
      if (first === undefined || firstIndex === index) return [];
      return [this._clash({ clash, named: repeat, otherPath: first.path })];
    });
  }
  private _clash({
    clash,
    named,
    otherPath,
  }: ClashProps): SyncClashingPlacements {
    return {
      kind: "clashingPlacements",
      clash,
      table: this.table,
      path: named.path,
      otherPath,
      column: named.column,
    };
  }
  private get _befores(): NamedAt[] {
    return (this.entry.moves ?? []).map(({ before }, index) => ({
      column: before,
      path: moveBeforePath(index),
    }));
  }
  private get _newColumnsBefores(): NamedAt[] {
    return newColumnsBeforesOf(this.entry).map(({ before, path }) => ({
      column: before,
      path,
    }));
  }
  private get _movedColumns(): NamedAt[] {
    return (this.entry.moves ?? []).flatMap(({ columns }, index) =>
      columns.map((column, columnIndex) => ({
        column,
        path: moveColumnPath({ index, columnIndex }),
      })),
    );
  }
}
