import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { ApiColumnIdentities } from "../ColumnConfigsOperator/ApiColumnIdentities";
import { InstalledBraceConfigs } from "../ColumnConfigsOperator/InstalledBraceConfigs";
import { LiveColumnFinder } from "../ColumnConfigsOperator/LiveColumnFinder";
import {
  type ConfigSyncState,
  SpreadsheetBaseOperator,
} from "../SpreadsheetBaseOperator";
import { TableConfigsOperator } from "../TableConfigsOperator";

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
  get liveColumnFinder(): LiveColumnFinder {
    return new LiveColumnFinder({
      ss: this.ss,
      apiColumnIdentities: this.apiColumnIdentities,
    });
  }
}
