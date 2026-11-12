import {
  makeImportLine,
  type TabConfigsBase,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { oneLinePerEntryFileSource } from "./configFileSource";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";

export class TabConfigOperator extends SpreadsheetBaseOperator {
  static init(): TabConfigOperator {
    return new TabConfigOperator(SpreadsheetBaseOperator.initOperatorProps());
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  newTabConfigs(): TabConfigsBase {
    const { raw } = this.ss;
    return raw.activeSheetGids.reduce<TabConfigsBase>(
      (tabConfigs, sheetGid) => {
        const sheet = raw.sheet(sheetGid);
        const tabName = this.schema.titleToName(sheet.title);
        const existing = tabConfigs[tabName];
        if (existing !== undefined) {
          throw new Error(
            `Tabs "${existing.title}" and "${sheet.title}" both give the tab name "${tabName}".`,
          );
        }
        const { color, ...tabProperties } = sheet.tabProperties;
        tabConfigs[tabName] = {
          sheetGid,
          title: sheet.title,
          color: color ?? null,
          ...tabProperties,
        };
        return tabConfigs;
      },
      {},
    );
  }
  toFileSource(): string {
    return [
      makeImportLine(["makeTabConfigs"]),
      ``,
      `export const tabConfigs = makeTabConfigs(${oneLinePerEntryFileSource(
        this.newTabConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
}
