import { Val } from "../../../utils/Val";
import {
  type EditProtection,
  type EditProtectionContent,
  isWholeSheetGridRange,
  type ProtectionGridRange,
} from "../../RawSource/EditProtection";
import type { GoogleRequest } from "../GoogleSheetsAPI";
import { googleGrid } from "./gridSnapshots";

type GoogleProtectedRange = GoogleAppsScript.Sheets.Schema.ProtectedRange;
type GoogleGridRange = GoogleAppsScript.Sheets.Schema.GridRange;
type BatchUpdateResponse =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetResponse;

export const googleProtectedRange = {
  fromContent(protection: EditProtectionContent): GoogleProtectedRange {
    const editors =
      protection.kind === "lock" &&
      (protection.users.length > 0 || protection.groups.length > 0)
        ? {
            ...(protection.users.length > 0 ? { users: protection.users } : {}),
            ...(protection.groups.length > 0
              ? { groups: protection.groups }
              : {}),
          }
        : undefined;
    return {
      range: googleProtectionRange.fromProtectionGridRange(protection.range),
      ...(protection.description !== ""
        ? { description: protection.description }
        : {}),
      ...(protection.kind === "warning" ? { warningOnly: true } : {}),
      ...(protection.unprotectedRanges.length > 0
        ? {
            unprotectedRanges: protection.unprotectedRanges.map(
              googleProtectionRange.fromProtectionGridRange,
            ),
          }
        : {}),
      ...(editors !== undefined ? { editors } : {}),
    };
  },
  toEditProtection(protection: GoogleProtectedRange): EditProtection {
    const id = Val.assert(protection.protectedRangeId, "protectedRangeId");
    if (protection.namedRangeId !== undefined) {
      return { kind: "unmodelable", id };
    }
    if (protection.range === undefined) {
      return { kind: "unmodelable", id };
    }
    if (protection.editors?.domainUsersCanEdit === true) {
      return { kind: "unmodelable", id };
    }
    const isWarning = protection.warningOnly === true;
    // Google lists editors on a warning too, but a warning never uses them.
    const editors = isWarning ? undefined : protection.editors;
    return {
      kind: isWarning ? "warning" : "lock",
      id,
      range: googleProtectionRange.toProtectionGridRange(protection.range),
      description: protection.description ?? "",
      users: editors?.users ?? [],
      groups: editors?.groups ?? [],
      unprotectedRanges: (protection.unprotectedRanges ?? []).map(
        googleProtectionRange.toProtectionGridRange,
      ),
      requestingUserCanEdit: protection.requestingUserCanEdit ?? false,
    };
  },
  validateAddReplies(
    response: BatchUpdateResponse,
    requests: GoogleRequest[],
  ): void {
    const replies = response.replies;
    if (replies === undefined) return;
    requests.forEach((request, index) => {
      if (request.addProtectedRange === undefined) return;
      const id =
        replies[index]?.addProtectedRange?.protectedRange?.protectedRangeId;
      if (id === undefined) {
        throw new Error(
          "Add protected range reply did not include a protectedRangeId.",
        );
      }
    });
  },
};

const googleProtectionRange = {
  fromProtectionGridRange(range: ProtectionGridRange): GoogleGridRange {
    if (isWholeSheetGridRange(range)) return { sheetId: range.sheetId };
    return range;
  },
  toProtectionGridRange(range: GoogleGridRange): ProtectionGridRange {
    const hasBound =
      range.startRowIndex !== undefined ||
      range.endRowIndex !== undefined ||
      range.startColumnIndex !== undefined ||
      range.endColumnIndex !== undefined;
    if (!hasBound) {
      return { sheetId: range.sheetId ?? 0 };
    }
    return googleGrid.toGridRangeProps(range);
  },
};
