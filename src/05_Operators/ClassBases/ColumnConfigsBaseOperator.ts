import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed.ts";
import { ApiColumnIdentities } from "../ColumnConfigsOperator/ApiColumnIdentities.ts";
import { InstalledBraceConfigs } from "../ColumnConfigsOperator/InstalledBraceConfigs.ts";
import {
  type ConfigSyncState,
  SpreadsheetBaseOperator,
} from "../SpreadsheetBaseOperator.ts";
import { TableConfigsOperator } from "../TableConfigsOperator.ts";

export class ColumnConfigsBaseOperator extends SpreadsheetBaseOperator {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get columnConfigSync(): ConfigSyncState["columnConfigSync"] {
    return this.configSyncState.columnConfigSync;
  }
  get tableConfigOperator(): TableConfigsOperator {
    return new TableConfigsOperator(this.operatorProps);
  }
  get apiColumnIdentities(): ApiColumnIdentities {
    return new ApiColumnIdentities({
      ss: this.ss,
      tableConfigOperator: this.tableConfigOperator,
    });
  }
  get installedBraceConfigs(): InstalledBraceConfigs {
    return new InstalledBraceConfigs();
  }
}
