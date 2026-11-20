import {
  entriesOf,
  movePath,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";

type ClashingMoves = Extract<SyncRefusal, { kind: "clashingMoves" }>;

interface NamedAt {
  column: string;
  path: string;
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
  private get _sameBeforeRefusals(): ClashingMoves[] {
    return this._repeatRefusals(this._befores, "sameBefore");
  }
  private get _repeatedColumnRefusals(): ClashingMoves[] {
    return this._repeatRefusals(this._movedColumns, "repeatedColumn");
  }
  // A group's target can't move, its own group's included, or one group's landing would undo another's.
  private get _movedBeforeRefusals(): ClashingMoves[] {
    const movedColumns = this._movedColumns;
    return this._befores.flatMap(({ column, path }) => {
      const moved = movedColumns.find((named) => named.column === column);
      if (moved === undefined) return [];
      return [this._clash("movedBefore", { column, path }, moved.path)];
    });
  }
  // A column placed twice can't sit in both places.
  private get _ownBeforeRefusals(): ClashingMoves[] {
    const ownBefores = new Set(
      entriesOf(this.entry.columns).flatMap(([column, { before }]) =>
        before === undefined ? [] : [column],
      ),
    );
    return this._movedColumns.flatMap((named) => {
      if (!ownBefores.has(named.column)) return [];
      return [this._clash("ownBefore", named, `${named.column}.before`)];
    });
  }
  private _repeatRefusals(
    named: NamedAt[],
    clash: ClashingMoves["clash"],
  ): ClashingMoves[] {
    return named.flatMap((repeat, index) => {
      const firstIndex = named.findIndex(
        ({ column }) => column === repeat.column,
      );
      const first = named[firstIndex];
      if (first === undefined || firstIndex === index) return [];
      return [this._clash(clash, repeat, first.path)];
    });
  }
  private _clash(
    clash: ClashingMoves["clash"],
    { column, path }: NamedAt,
    otherPath: string,
  ): ClashingMoves {
    return {
      kind: "clashingMoves",
      clash,
      table: this.table,
      path,
      otherPath,
      column,
    };
  }
  private get _befores(): NamedAt[] {
    return (this.entry.moves ?? []).map(({ before }, index) => ({
      column: before,
      path: `${movePath(index)}.before`,
    }));
  }
  private get _movedColumns(): NamedAt[] {
    return (this.entry.moves ?? []).flatMap(({ columns }, index) =>
      columns.map((column, columnIndex) => ({
        column,
        path: `${movePath(index)}.columns[${columnIndex}]`,
      })),
    );
  }
}
