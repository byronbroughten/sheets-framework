import {
  type EditLockDeclaration,
  type EditProtection,
  type EditProtectionContent,
  editProtectionContentSatisfies,
  editProtectionsEqual,
  type EditWarningDeclaration,
  isWholeColumnGridRange,
  type ProtectionGridRange,
  protectionRangeEqual,
  protectionRangeHasRowCoordinates,
  type WholeSheetEditLockDeclaration,
  type WholeSheetEditWarningDeclaration,
} from "../../00_Source/RawSource/EditProtection";
import { SheetCommonRaw } from "../ClassBases/SheetCommonRaw";
import { SheetRaw } from "../SheetRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

export class SheetEditProtectionsRaw extends SheetCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get sheet(): SheetRaw {
    return new SheetRaw(this.sheetRawProps);
  }
  gatherFetchEditProtections(): void {
    this.sheetState.fetchQueue.gatherEditProtections = true;
  }
  editProtections(): EditProtection[] {
    this.assertEditProtectionsNotStale();
    const protections = this.sheetState.working.editProtections.protections;
    if (protections === undefined) {
      throw new Error(
        `Edit protections have not been fetched for sheetGid ${this.sheetGid}.`,
      );
    }
    return protections;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): void {
    this.addEditWarningAt(this.sheet.dataGridRange, declaration);
  }
  addEditLock(declaration: EditLockDeclaration = {}): void {
    this.addEditLockAt(this.sheet.dataGridRange, declaration);
  }
  addEditWarningWholeSheet(
    declaration: WholeSheetEditWarningDeclaration = {},
  ): void {
    this._queueProtection({
      kind: "warning",
      range: this.sheet.wholeSheetGridRange,
      description: declaration.description ?? "",
      users: [],
      groups: [],
      unprotectedRanges: declaration.unprotectedRanges ?? [],
    });
  }
  addEditLockWholeSheet(declaration: WholeSheetEditLockDeclaration = {}): void {
    this._queueProtection({
      kind: "lock",
      range: this.sheet.wholeSheetGridRange,
      description: declaration.description ?? "",
      users: declaration.users ?? [],
      groups: declaration.groups ?? [],
      unprotectedRanges: declaration.unprotectedRanges ?? [],
    });
  }
  addEditWarningAt(
    range: ProtectionGridRange,
    declaration: EditWarningDeclaration = {},
  ): void {
    this._queueProtection({
      kind: "warning",
      range,
      description: declaration.description ?? "",
      users: [],
      groups: [],
      unprotectedRanges: [],
    });
  }
  addEditLockAt(
    range: ProtectionGridRange,
    declaration: EditLockDeclaration = {},
  ): void {
    this._queueProtection({
      kind: "lock",
      range,
      description: declaration.description ?? "",
      users: declaration.users ?? [],
      groups: declaration.groups ?? [],
      unprotectedRanges: [],
    });
  }
  private _queueProtection(protection: EditProtectionContent): void {
    this._assertProtectionWriteCoordinatesNotStale(
      protection.range,
      protection.unprotectedRanges,
    );
    this.assertEditProtectionsNotStale();
    if (
      this._pendingEditProtectionContents().some((pending) =>
        editProtectionContentSatisfies(pending, protection),
      )
    ) {
      return;
    }
    this.updateRequests.addProtectedRange.push({
      kind: "addProtectedRange",
      protection,
    });
  }
  private _pendingEditProtectionContents(): EditProtectionContent[] {
    const fetched = this.sheetState.working.editProtections.protections;
    const protections: EditProtection[] =
      fetched === undefined ? [] : [...fetched];
    const deletedIds = new Set(
      this.updateRequests.deleteProtectedRange
        .filter((operation) => operation.sheetId === this.sheetGid)
        .map((operation) => operation.protectedRangeId),
    );
    const remaining = protections.filter(
      (protection) => !deletedIds.has(protection.id),
    );
    const queued = this.updateRequests.addProtectedRange
      .filter(
        (operation) => operation.protection.range.sheetId === this.sheetGid,
      )
      .map((operation) => operation.protection);
    return [
      ...remaining.flatMap((protection) =>
        protection.kind === "unmodelable" ? [] : [protection],
      ),
      ...queued,
    ];
  }
  removeEditProtections(): void {
    this.removeEditProtectionsAt(this.sheet.dataGridRange);
  }
  removeEditProtectionsAt(range: ProtectionGridRange): void {
    this._assertProtectionWriteCoordinatesNotStale(range, []);
    this.assertEditProtectionsNotStale();
    this.editProtections().forEach((protection) => {
      if (protection.kind === "unmodelable") return;
      if (!protectionRangeEqual(range, protection.range)) return;
      this._queueDeleteEditProtection(protection.id);
    });
  }
  removeEditProtection(protection: EditProtection): void {
    this._removeProtectionsWhere((existing) =>
      editProtectionsEqual(existing, protection),
    );
  }
  removeEditProtectionByDescription(description: string): void {
    this._removeProtectionsWhere(
      (existing) =>
        existing.kind !== "unmodelable" && existing.description === description,
    );
  }
  removeEditProtectionById(protectionId: number): void {
    this._removeProtectionsWhere((existing) => existing.id === protectionId);
  }
  private _removeProtectionsWhere(
    matches: (protection: EditProtection) => boolean,
  ): void {
    this.assertEditProtectionsNotStale();
    this.editProtections().forEach((existing) => {
      if (!matches(existing)) return;
      if (existing.kind !== "unmodelable") {
        this._assertProtectionWriteCoordinatesNotStale(
          existing.range,
          existing.unprotectedRanges,
        );
      }
      this._queueDeleteEditProtection(existing.id);
    });
  }
  private _queueDeleteEditProtection(protectionId: number): void {
    this.updateRequests.deleteProtectedRange.push({
      kind: "deleteProtectedRange",
      sheetId: this.sheetGid,
      protectedRangeId: protectionId,
    });
  }
  private _assertProtectionWriteCoordinatesNotStale(
    range: ProtectionGridRange,
    unprotectedRanges: ProtectionGridRange[],
  ): void {
    [range, ...unprotectedRanges].forEach((item) => {
      this._assertOneProtectionRangeCoordinatesNotStale(item);
    });
  }
  private _assertOneProtectionRangeCoordinatesNotStale(
    range: ProtectionGridRange,
  ): void {
    if (isWholeColumnGridRange(range)) {
      if (this.sheetState.working.knownTable !== undefined) {
        this.activeTable.validateColIndexNotStale(range.startColumnIndex);
      }
      return;
    }
    if (!protectionRangeHasRowCoordinates(range)) return;
    this.activeTable.assertRowIndexesNotStale();
  }
  markEditProtectionsStale(): void {
    this.sheetState.working.editProtections.isStale = true;
  }
  assertEditProtectionsNotStale(): void {
    if (!this.sheetState.working.editProtections.isStale) return;
    throw new Error(
      `Edit protections are stale for sheetGid ${this.sheetGid}. Re-fetch the sheet's protections before reading or mutating them again.`,
    );
  }
  integrateEditProtections(protections: EditProtection[]): void {
    this.sheetState.working.editProtections.protections = protections;
    this.sheetState.fetchQueue.gatherEditProtections = false;
    this.sheetState.working.editProtections.isStale = false;
  }
}
