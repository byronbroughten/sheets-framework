import {
  type IntendedTrait,
  isTableDeletion,
} from "../../01_SpreadsheetSchema/configReaders/SyncPlanner.ts";
import { previousTableIdOf } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes.ts";
import type { TableRaw } from "../../02_SpreadsheetRaw/TableRaw.ts";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed.ts";
import { SpreadsheetBaseOperator } from "../SpreadsheetBaseOperator.ts";

export class TableConfigsTraitUpdaterOperator extends SpreadsheetBaseOperator {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  // A renamed Table is addressed by its previous key, which the refuser guarantees the previous configs hold.
  updateTableNames(sent: IntendedTrait[]): this {
    for (const intended of sent) {
      if (intended.kind !== "table" || intended.trait !== "name") continue;
      this._previousTable(intended.table).updateTableName(
        String(intended.value),
      );
    }
    return this;
  }
  deleteTables(sent: IntendedTrait[]): this {
    for (const intended of sent) {
      if (intended.kind !== "table" || !isTableDeletion(intended)) continue;
      this._previousTable(intended.table).queueDelete();
    }
    return this;
  }
  private _previousTable(tableKey: string): TableRaw {
    return this.ss.raw.table(previousTableIdOf(tableKey));
  }
}
