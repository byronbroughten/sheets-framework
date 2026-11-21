import { Val } from "@byronbroughten/utils/val";

import type { SheetTabProperties } from "../00_Source/RawSource/RawSource";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import {
  makeImportLine,
  type TabConfigsBase,
  type TabConfigStored,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  type IntendedTrait,
  isTabDeletion,
  type TraitReadings,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type { SyncTabGoneNotice } from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { oneLinePerEntryFileSource } from "./configFileSource";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";

type TabReadings = TraitReadings["tabs"];
type TabReading = NonNullable<TabReadings[string]>;

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
  // `before` is read live only: order is never in the configs, so it can't conflict.
  traitReadings(): { previous: TabReadings; live: TabReadings } {
    const tabNames = [...this._sheetsByTabName().keys()];
    const live = Object.entries(this.newTabConfigs()).reduce<TabReadings>(
      (readings, [tabName, config]) => {
        const next = tabNames[tabNames.indexOf(tabName) + 1];
        readings[tabName] = {
          ...tabTraitsOf(config),
          ...(next !== undefined ? { before: next } : {}),
        };
        return readings;
      },
      {},
    );
    const previous = Object.entries(
      installedConfigs().tabConfigs ?? {},
    ).reduce<TabReadings>((readings, [tabName, config]) => {
      readings[tabName] = tabTraitsOf(config);
      return readings;
    }, {});
    return { previous, live };
  }
  // Tabs resolve by their names before any retitle, which are the names the readings use.
  updateTabTraits(sent: IntendedTrait[]): this {
    const sheets = this._sheetsByTabName();
    return this._updateTabTraits(sent, (tabName) =>
      liveSheetOf(sheets, tabName),
    );
  }
  // After the refetch a retitled tab has its new title, so a tab the previous map holds resolves by its GID and a new tab by its title.
  updateNewTabTraits(
    newTabTraits: IntendedTrait[],
    { tabs }: IdNameMapGeneric,
  ): this {
    const sheets = this._sheetsByTabName();
    const gidsByTabName = new Map(
      Object.entries(tabs).map(([sheetGid, { name }]) => [
        name,
        Number(sheetGid),
      ]),
    );
    const { raw } = this.ss;
    return this._updateTabTraits(newTabTraits, (tabName) => {
      const sheetGid = gidsByTabName.get(tabName);
      // A tab a person deleted keeps its map entry, and raw.sheet would build it anyway.
      if (sheetGid !== undefined && raw.gidIsActive(sheetGid)) {
        return raw.sheet(sheetGid);
      }
      return liveSheetOf(sheets, tabName);
    });
  }
  private _updateTabTraits(
    sent: IntendedTrait[],
    sheetNamed: (tabName: string) => SheetRaw,
  ): this {
    // A deleted tab's other traits would be lost with it, and a retitle would hide it from deleteTabs.
    const deletedTabs = deletedTabsOf(sent);
    for (const intended of sent) {
      if (intended.kind !== "tab" || deletedTabs.has(intended.tab)) continue;
      const sheet = sheetNamed(intended.tab);
      const { trait, value } = intended;
      if (trait === "title") {
        sheet.updateTitle(String(value));
      } else if (trait === "before") {
        sheet.moveBefore(sheetNamed(String(value)));
      } else {
        // The planner picks only SheetTabProperties keys besides title and before.
        sheet.updateTabProperties({
          [trait]: value,
        } as Partial<SheetTabProperties>);
      }
    }
    return this;
  }
  // By the names before any retitle, which updateTabTraits leaves a deleted tab without.
  deleteTabs(sent: IntendedTrait[]): this {
    const sheets = this._sheetsByTabName();
    for (const tab of deletedTabsOf(sent)) {
      liveSheetOf(sheets, tab).queueDelete();
    }
    return this;
  }
  // Read before any send, so a tab this run deletes is still live and only a person's deletion shows.
  goneTabNotices({ tabs }: IdNameMapGeneric): SyncTabGoneNotice[] {
    const liveGids = new Set(this.ss.raw.activeSheetGids.map(String));
    return Object.entries(tabs).flatMap(
      ([sheetGid, { name }]): SyncTabGoneNotice[] =>
        liveGids.has(sheetGid) ? [] : [{ kind: "tabGone", tab: name }],
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
  private _sheetsByTabName(): Map<string, SheetRaw> {
    const { raw } = this.ss;
    return new Map(
      raw.activeSheetGids.map((sheetGid) => {
        const sheet = raw.sheet(sheetGid);
        return [this.schema.titleToName(sheet.title), sheet];
      }),
    );
  }
}

function liveSheetOf(sheets: Map<string, SheetRaw>, tab: string): SheetRaw {
  return Val.assert(sheets.get(tab), `live tab ${tab}`);
}

function deletedTabsOf(sent: IntendedTrait[]): Set<string> {
  return new Set(
    sent.flatMap((intended) => {
      if (!isTabDeletion(intended)) return [];
      return [intended.tab];
    }),
  );
}

// The configs write no colour as null; the planner's absent value is undefined.
function tabTraitsOf({
  title,
  color,
  hidden,
  frozenRows,
  frozenColumns,
}: TabConfigStored): TabReading {
  return {
    title,
    ...(color !== null ? { color } : {}),
    hidden,
    frozenRows,
    frozenColumns,
  };
}
