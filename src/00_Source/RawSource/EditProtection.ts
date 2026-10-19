import {
  type GridRangeProps,
  isWholeSheetGridRange,
  rangeEqual,
  type WholeSheetGridRange,
} from "./RawSource";

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
    contentsEqualBesidesEditors(left, right) &&
    stringListsEqual(left.users, right.users) &&
    stringListsEqual(left.groups, right.groups)
  );
}

// Google adds its own editors to a lock, so a present lock need only include the declared ones.
export function editProtectionContentSatisfies(
  present: EditProtectionContent,
  declared: EditProtectionContent,
): boolean {
  return (
    contentsEqualBesidesEditors(present, declared) &&
    stringListIncludesAll(present.users, declared.users) &&
    stringListIncludesAll(present.groups, declared.groups)
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

function contentsEqualBesidesEditors(
  left: EditProtectionContent,
  right: EditProtectionContent,
): boolean {
  return (
    left.kind === right.kind &&
    protectionRangeEqual(left.range, right.range) &&
    left.description === right.description &&
    protectionRangesEqual(left.unprotectedRanges, right.unprotectedRanges)
  );
}

function stringListIncludesAll(list: string[], required: string[]): boolean {
  return required.every((value) => list.includes(value));
}

function stringListsEqual(left: string[], right: string[]): boolean {
  if (left.length !== right.length) return false;
  return left.every((value, index) => value === right[index]);
}
