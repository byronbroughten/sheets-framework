import {
  entriesOf,
  moveBeforePath,
  moveColumnPath,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncClashingMoves, SyncRefusal } from "../syncRunResult";

interface NamedAt {
  column: string;
  path: string;
}

interface ClashProps {
  clash: SyncClashingMoves["clash"];
  named: NamedAt;
  otherPath: string;
}

// Refuses a Table entry's `moves` groups that can't all hold at once, naming both paths of each clash.
export class MovesRefuser {
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
      ...this._ownBeforeRefusals,
    ];
  }
  private get _sameBeforeRefusals(): SyncClashingMoves[] {
    return this._repeatRefusals(this._befores, "sameBefore");
  }
  private get _repeatedColumnRefusals(): SyncClashingMoves[] {
    return this._repeatRefusals(this._movedColumns, "repeatedColumn");
  }
  // A group's target can't move, its own group's included, or one group's landing would undo another's.
  private get _movedBeforeRefusals(): SyncClashingMoves[] {
    const movedColumns = this._movedColumns;
    return this._befores.flatMap((named) => {
      const moved = movedColumns.find(({ column }) => column === named.column);
      if (moved === undefined) return [];
      return [
        this._clash({ clash: "movedBefore", named, otherPath: moved.path }),
      ];
    });
  }
  // A column placed twice can't sit in both places.
  private get _ownBeforeRefusals(): SyncClashingMoves[] {
    const ownBefores = new Set(
      entriesOf(this.entry.columns).flatMap(([column, { before }]) =>
        before === undefined ? [] : [column],
      ),
    );
    return this._movedColumns.flatMap((named) => {
      if (!ownBefores.has(named.column)) return [];
      return [
        this._clash({
          clash: "ownBefore",
          named,
          otherPath: `${named.column}.before`,
        }),
      ];
    });
  }
  private _repeatRefusals(
    named: NamedAt[],
    clash: SyncClashingMoves["clash"],
  ): SyncClashingMoves[] {
    return named.flatMap((repeat, index) => {
      const firstIndex = named.findIndex(
        ({ column }) => column === repeat.column,
      );
      const first = named[firstIndex];
      if (first === undefined || firstIndex === index) return [];
      return [this._clash({ clash, named: repeat, otherPath: first.path })];
    });
  }
  private _clash({ clash, named, otherPath }: ClashProps): SyncClashingMoves {
    return {
      kind: "clashingMoves",
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
  private get _movedColumns(): NamedAt[] {
    return (this.entry.moves ?? []).flatMap(({ columns }, index) =>
      columns.map((column, columnIndex) => ({
        column,
        path: moveColumnPath({ index, columnIndex }),
      })),
    );
  }
}
