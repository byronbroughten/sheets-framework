import {
  type ModelableEditProtection,
  protectionRangeEqual,
  protectionRangesEqual,
} from "../../00_Source/RawSource/EditProtection";
import { getSheetTraitByName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SpreadsheetBaseNamed } from "../../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { type FloorSheetName, floorSheetNames } from "./floorSeedLookups";
import {
  type FloorDeclaration,
  FloorTabEditWarning,
  floorWarningPrefix,
} from "./FloorTabEditWarning";

// Sheet indexes of the columns that say whose row it is, gathered with the floor's fetch.
export type IdentityColIndexes = Map<FloorSheetName, number[]>;

interface FloorTabDeclaration {
  tab: FloorTabEditWarning<FloorSheetName>;
  declaration: FloorDeclaration;
}

/**
 * Reconciles the floor's edit warnings across the floor tabs: Spreadsheet
 * Config, Sheet Config and Column Config each get one, and any that drifted
 * or no floor tab declares is removed. ConfigSheetFloor runs this after its restores.
 * Each tab's declaration and the rules table live in FloorTabEditWarning; the
 * live column lookup is floorColumnLocation; seed lookups are floorSeedLookups;
 * ConfigSheetFloorCreator creates missing floor tabs and columns before this runs.
 * docs/generated-data/config-sheet-floor.md
 */
export class ConfigSheetFloorEditWarnings extends SpreadsheetBaseNamed {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  gatherIdentityColumns(): IdentityColIndexes {
    const identityColIndexes: IdentityColIndexes = new Map();
    floorSheetNames().forEach((sheetName) => {
      const colIndexes = this._floorTab(sheetName).gatherIdentityColumns();
      if (colIndexes !== undefined) {
        identityColIndexes.set(sheetName, colIndexes);
      }
    });
    return identityColIndexes;
  }
  ensure(identityColIndexes: IdentityColIndexes): string[] {
    const report: string[] = [];
    const tabs = this._activeFloorSheetNames().map((sheetName) =>
      this._floorTab(sheetName),
    );
    tabs.forEach((tab) => {
      const addedColumnLines = tab.addedColumnReportLines();
      if (addedColumnLines.length > 0) {
        report.push(`Covered added columns: ${addedColumnLines.join("; ")}`);
      }
    });
    const tabDeclarations = tabs.map((tab) => ({
      tab,
      declaration: tab.declaration(identityColIndexes.get(tab.sheetName)),
    }));
    this._reconcile(tabDeclarations, report);
    return report;
  }
  private _floorTab(
    sheetName: FloorSheetName,
  ): FloorTabEditWarning<FloorSheetName> {
    return new FloorTabEditWarning({
      ...this.spreadsheetNamedProps,
      sheetName,
    });
  }
  private _reconcile(
    tabDeclarations: FloorTabDeclaration[],
    report: string[],
  ): void {
    const existing = this._floorProtections();
    const claimed = new Set<number>();
    const removed = new Set<number>();
    const drifted: string[] = [];
    tabDeclarations.forEach(({ tab, declaration }) => {
      const matches = existing.filter(
        (protection) => protection.description === declaration.description,
      );
      const inPlace = matches.find(
        (protection) =>
          protectionRangeEqual(protection.range, declaration.range) &&
          protectionRangesEqual(
            protection.unprotectedRanges,
            declaration.unprotectedRanges,
          ),
      );
      if (inPlace !== undefined) {
        claimed.add(inPlace.id);
        return;
      }
      matches.forEach((protection) => {
        this._removeProtection(protection);
        removed.add(protection.id);
        drifted.push(protection.description);
      });
      tab.queueAdd(declaration);
    });
    existing.forEach((protection) => {
      if (claimed.has(protection.id) || removed.has(protection.id)) return;
      this._removeProtection(protection);
    });
    if (drifted.length > 0) {
      report.push(`Replaced drifted: ${drifted.join("; ")}`);
    }
  }
  // A floor tab created this run is still absent after the refetch under a fake or dry run.
  private _activeFloorSheetNames(): FloorSheetName[] {
    return floorSheetNames().filter((sheetName) =>
      this.ss.raw.gidIsActive(getSheetTraitByName(sheetName, "sheetGid")),
    );
  }
  private _floorProtections(): ModelableEditProtection[] {
    return this._activeFloorSheetNames().flatMap((sheetName) =>
      this.ss
        .sheet(sheetName)
        .editProtections()
        .flatMap((protection) => {
          if (protection.kind === "unmodelable") return [];
          if (!protection.description.startsWith(floorWarningPrefix)) {
            return [];
          }
          return [protection];
        }),
    );
  }
  private _removeProtection(protection: ModelableEditProtection): void {
    floorSheetNames().forEach((sheetName) => {
      const sheet = this.ss.sheet(sheetName);
      if (sheet.schema.sheetGid !== protection.range.sheetId) return;
      sheet.removeEditProtectionById(protection.id);
    });
  }
}
