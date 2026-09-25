import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import {
  configSheetFloorSeed,
  floorTabSeedByGid,
} from "../01_SpreadsheetSchema/configSheetFloorSeed";
import {
  idPrefixes,
  type IdPrefixLabel,
} from "../01_SpreadsheetSchema/idPrefixes";
import {
  makeImportLine,
  type SheetConfigsBase,
} from "../01_SpreadsheetSchema/makeConfigs";
import { sheetConfigsByGid } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { Val } from "../utils/Val";
import { sheetConfigsFileSource } from "./configFileSource";
import { GenericSheetOperator } from "./GenericSheetOperator";
import {
  type ConfigSyncState,
  type OperatorProps,
  SpreadsheetBaseOperator,
} from "./SpreadsheetBaseOperator";

export class SheetConfigOperator extends GenericSheetOperator<"sheetConfig"> {
  constructor(props: OperatorProps) {
    super({
      sheetName: "sheetConfig",
      ...props,
    });
  }
  static init(): SheetConfigOperator {
    return new SheetConfigOperator(SpreadsheetBaseOperator.initOperatorProps());
  }
  get sheetConfigSync(): ConfigSyncState["sheetConfigSync"] {
    return this.configSyncState.sheetConfigSync;
  }
  assertPrepFetchIsComplete(): void {
    if (!this.sheetConfigSync.prepFetchIsComplete) {
      throw new Error(
        "SheetConfigOperator has not yet completed its prepFetch operation.",
      );
    }
  }
  assertSyncedToSpreadsheet(): void {
    if (!this.sheetConfigSync.syncedToSpreadsheet) {
      throw new Error(
        "SheetConfigOperator has not yet synced to the spreadsheet.",
      );
    }
  }
  prepFetchForSync(): void {
    this.sheet.prepFetchColumnsFull("sheetGid", "sheetTitle", "letApiAccess");
    this.sheetConfigSync.prepFetchIsComplete = true;
  }
  syncToSpreadsheet(): void {
    this._deleteStaleSheetConfigs();
    this._appendMissingSheetConfigs();
    this._updateProgrammaticValues();
    this.sheetConfigSync.syncedToSpreadsheet = true;
  }
  private _deleteStaleSheetConfigs(): void {
    this.sheet.rows.forEach((row) => {
      const configGid = row.valueOrEmpty("sheetGid");
      if (configGid === "" || !this.ss.raw.gidIsActive(configGid)) {
        row.delete();
      }
    });
  }
  private _appendMissingSheetConfigs(): void {
    const colGid = this.sheet.column("sheetGid");
    this.ss.raw.activeSheetGids.forEach((sheetGid) => {
      if (!colGid.hasValue(sheetGid)) {
        this.sheet.appendRowWithVals({ sheetGid });
      }
    });
  }
  private _updateProgrammaticValues(): void {
    const col = this.sheet.columns("sheetGid", "sheetTitle", "letApiAccess");
    const reportLines = this.sheetConfigSync.declaredCellReportLines;
    reportLines.length = 0;
    let updatedValues = 0;
    this.sheet.rowIndexesActiveWithData.forEach((rowIndex) => {
      const sheetGid = col.sheetGid.value(rowIndex);
      const activeSheet = this.ss.raw.sheet(sheetGid);
      if (col.sheetTitle.valueOrEmpty(rowIndex) !== activeSheet.title) {
        col.sheetTitle.cell(rowIndex).updateValue(activeSheet.title);
        updatedValues++;
      }
      if (
        this._updateSelfDescribingLetApiAccess({
          rowIndex,
          sheetGid,
          sheetTitle: activeSheet.title,
        })
      ) {
        updatedValues++;
      }
    });
    Logger.log(`Corrected ${updatedValues} inaccurate Sheet Config cells.`);
  }
  private _updateSelfDescribingLetApiAccess({
    rowIndex,
    sheetGid,
    sheetTitle,
  }: {
    rowIndex: number;
    sheetGid: number;
    sheetTitle: string;
  }): boolean {
    const seed = floorTabSeedByGid(sheetGid);
    if (seed === undefined) return false;
    const letApiAccess = this.sheet.column("letApiAccess");
    if (letApiAccess.valueOrEmpty(rowIndex) === seed.letApiAccess) return false;
    letApiAccess.cell(rowIndex).updateValue(seed.letApiAccess);
    this.sheetConfigSync.declaredCellReportLines.push(
      `${configSheetFloorSeed.sheetConfig.title} · ${getColumnTraitByName(
        "sheetConfig",
        "letApiAccess",
        "header",
      )} · ${sheetTitle} → ${seed.letApiAccess ? "TRUE" : "FALSE"}`,
    );
    return true;
  }
  isSheetGidApiAccess(sheetGid: number): boolean {
    return this.sheetGidsApiAccesses().includes(sheetGid);
  }
  sheetGidsApiAccesses(): number[] {
    const col = this.sheet.columns("sheetGid", "letApiAccess");
    const gids: number[] = [];
    this.sheet.rowIndexesActiveWithData.forEach((rowIndex) => {
      if (col.letApiAccess.valueOrEmpty(rowIndex)) {
        gids.push(col.sheetGid.value(rowIndex));
      }
    });
    return gids;
  }
  idPrefix(sheetGid: number): string {
    return Val.assert(this._idPrefixesBySheetGid().get(sheetGid), "ID prefix");
  }
  private _idPrefixesBySheetGid(): Map<number, string> {
    const prefixesInUse = new Set<string>();
    const assigned = new Map<number, string>();
    this.sheetGidsApiAccesses().forEach((sheetGid) => {
      const sampled = this.ss.raw.sheetMeta(sheetGid).activeIdPrefix();
      if (sampled === undefined) return;
      prefixesInUse.add(sampled);
      assigned.set(sheetGid, sampled);
    });
    this.sheetGidsApiAccesses().forEach((sheetGid) => {
      if (assigned.has(sheetGid)) return;
      const generated = idPrefixes.fromTitle(
        this.ss.raw.sheet(sheetGid).title,
        prefixesInUse,
      );
      prefixesInUse.add(generated);
      assigned.set(sheetGid, generated);
    });
    return assigned;
  }
  idPrefixChangeReport(): string | undefined {
    const changes: string[] = [];
    this.sheetGidsApiAccesses().forEach((sheetGid) => {
      const previous = sheetConfigsByGid().get(sheetGid);
      if (previous === undefined) return;
      const sampled = this.ss.raw.sheetMeta(sheetGid).activeIdPrefix();
      if (sampled === undefined || sampled === previous.idPrefix) return;
      changes.push(
        `Sheet "${this.ss.raw.sheet(sheetGid).title}" sampled ID prefix "${sampled}" differs from last generated "${previous.idPrefix}".`,
      );
    });
    if (changes.length === 0) return undefined;
    return changes.join(" ");
  }
  declaredCellReport(): string | undefined {
    const lines = this.sheetConfigSync.declaredCellReportLines;
    if (lines.length === 0) return undefined;
    return lines.join(" ");
  }
  newSheetConfigs(): SheetConfigsBase {
    const col = this.sheet.columns("sheetGid", "sheetTitle", "letApiAccess");
    const sheetConfigs: SheetConfigsBase = {};
    const idPrefixLabels: IdPrefixLabel[] = [];
    this.sheet.rowIndexesActiveWithData.forEach((rowIndex) => {
      // Defaults false on a freshly-appended row — excluded until a human sets it true in the sheet.
      if (!col.letApiAccess.valueOrEmpty(rowIndex)) return;
      const title = col.sheetTitle.value(rowIndex);
      const sheetName = this.schema.titleToName(title);
      const sheetGid = col.sheetGid.value(rowIndex);
      const idPrefix = this.idPrefix(sheetGid);
      const { tableHeaderRow } = this.ss.raw.sheetMeta(sheetGid);
      sheetConfigs[sheetName] = {
        sheetGid,
        idPrefix,
        hasIdColumn: tableHeaderRow.hasValue(this.schema.idHeader),
        hasNameColumn: tableHeaderRow.hasValue(this.schema.nameHeader),
      };
      idPrefixLabels.push({ label: title, idPrefix });
    });
    idPrefixes.assertUnique(idPrefixLabels);
    return sheetConfigs;
  }
  sheetNamesByGid(): Map<number, string> {
    const map = new Map<number, string>();
    Object.entries(this.newSheetConfigs()).forEach(([sheetName, config]) => {
      map.set(config.sheetGid, sheetName);
    });
    return map;
  }
  toFileSource(makeConfigsImport: string): string {
    return [
      `${makeImportLine("makeSheetConfigs", makeConfigsImport)}`,
      ``,
      `export const sheetConfigs = makeSheetConfigs(${sheetConfigsFileSource(
        this.newSheetConfigs(),
      )});`,
      ``,
    ].join("\n");
  }
}
