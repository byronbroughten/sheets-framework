import { rangeEqual } from "./ConditionalFormat";
import type { GridRangeProps } from "./RawSource";

export interface WholeSheetGridRange {
  sheetId: number;
}

export type ProtectionGridRange = GridRangeProps | WholeSheetGridRange;

export interface EditWarningDeclaration {
  description?: string;
}

export interface EditLockDeclaration {
  description?: string;
  users?: string[];
  groups?: string[];
}

export interface WholeSheetEditWarningDeclaration extends EditWarningDeclaration {
  unprotectedRanges?: ProtectionGridRange[];
}

export interface WholeSheetEditLockDeclaration extends EditLockDeclaration {
  unprotectedRanges?: ProtectionGridRange[];
}

export type EditProtection =
  ModelableEditProtection | UnmodelableEditProtection;

export interface ModelableEditProtection {
  kind: "warning" | "lock";
  id: number;
  range: ProtectionGridRange;
  description: string;
  users: string[];
  groups: string[];
  unprotectedRanges: ProtectionGridRange[];
  requestingUserCanEdit: boolean;
}

export interface UnmodelableEditProtection {
  kind: "unmodelable";
  id: number;
}

export type EditProtectionContent = Omit<
  ModelableEditProtection,
  "id" | "requestingUserCanEdit"
>;

export function isWholeSheetGridRange(
  range: ProtectionGridRange,
): range is WholeSheetGridRange {
  return !("startRowIndex" in range);
}

export function isWholeColumnGridRange(
  range: ProtectionGridRange,
): range is GridRangeProps & { startColumnIndex: number } {
  return (
    "startRowIndex" in range &&
    range.startRowIndex === 0 &&
    range.endRowIndex === undefined &&
    range.startColumnIndex !== undefined
  );
}

export function protectionRangeHasRowCoordinates(
  range: ProtectionGridRange,
): boolean {
  return "startRowIndex" in range || "endRowIndex" in range;
}

export function protectionRangeEqual(
  left: ProtectionGridRange,
  right: ProtectionGridRange | undefined,
): boolean {
  if (right === undefined) return false;
  if (isWholeSheetGridRange(left) || isWholeSheetGridRange(right)) {
    return (
      isWholeSheetGridRange(left) &&
      isWholeSheetGridRange(right) &&
      left.sheetId === right.sheetId
    );
  }
  return rangeEqual(left, right);
}

export function editProtectionContentsEqual(
  left: EditProtectionContent,
  right: EditProtectionContent,
): boolean {
  return (
    left.kind === right.kind &&
    protectionRangeEqual(left.range, right.range) &&
    left.description === right.description &&
    stringListsEqual(left.users, right.users) &&
    stringListsEqual(left.groups, right.groups) &&
    protectionRangesEqual(left.unprotectedRanges, right.unprotectedRanges)
  );
}

// Google adds its own editors to a lock, so a present lock need only include the declared ones.
export function editProtectionContentSatisfies(
  present: EditProtectionContent,
  declared: EditProtectionContent,
): boolean {
  return (
    present.kind === declared.kind &&
    protectionRangeEqual(present.range, declared.range) &&
    present.description === declared.description &&
    stringListIncludesAll(present.users, declared.users) &&
    stringListIncludesAll(present.groups, declared.groups) &&
    protectionRangesEqual(present.unprotectedRanges, declared.unprotectedRanges)
  );
}

export function editProtectionsEqual(
  left: EditProtection,
  right: EditProtection,
): boolean {
  if (left.kind === "unmodelable" || right.kind === "unmodelable") return false;
  return editProtectionContentsEqual(left, right);
}

export function protectionRangesEqual(
  left: ProtectionGridRange[],
  right: ProtectionGridRange[],
): boolean {
  if (left.length !== right.length) return false;
  return left.every((range, index) =>
    protectionRangeEqual(range, right[index]),
  );
}

function stringListIncludesAll(list: string[], required: string[]): boolean {
  return required.every((value) => list.includes(value));
}

function stringListsEqual(left: string[], right: string[]): boolean {
  if (left.length !== right.length) return false;
  return left.every((value, index) => value === right[index]);
}
