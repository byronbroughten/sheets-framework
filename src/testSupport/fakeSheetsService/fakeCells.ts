import type {
  FakeCell,
  FakeCellValue,
  FakeRichCellValue,
} from "../fakeSheetsService";

type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;
type ExtendedValue = GoogleAppsScript.Sheets.Schema.ExtendedValue;

// The cell-data field paths a write may name, and the cell fact each one replaces.
const cellFieldPaths = {
  userEnteredValue: ["value"],
  userEnteredFormat: ["backgroundColor", "numberFormat"],
  "userEnteredFormat.backgroundColor": ["backgroundColor"],
  "userEnteredFormat.numberFormat": ["numberFormat"],
  dataValidation: ["dataValidation"],
} as const;

type CellFact = (typeof cellFieldPaths)[keyof typeof cellFieldPaths][number];

export type CellFields = ReadonlySet<CellFact>;

export const fakeCells = {
  value(cell: FakeCell): FakeCellValue {
    return typeof cell === "object" && cell !== null ? cell.value : cell;
  },
  isBlank(cell: FakeCell): boolean {
    const value = fakeCells.value(cell);
    return value === null || value === "";
  },
  // Throws on a path the replay doesn't model, so a write can't pass by changing nothing.
  fields(fieldMask: string | undefined, requestKind: string): CellFields {
    if (fieldMask === undefined || fieldMask === "") {
      throw new Error(`${requestKind} needs a fields mask.`);
    }
    return new Set(
      fieldMask.split(",").flatMap((rawPath) => {
        const path = rawPath.trim();
        if (!isCellFieldPath(path)) {
          throw new Error(
            `The fake Sheets service does not replay ${requestKind} field "${path}".`,
          );
        }
        return [...cellFieldPaths[path]];
      }),
    );
  },
  // Replaces only the facts the mask names, as the live API does.
  withCellData(
    cell: FakeCell,
    data: GoogleCellData,
    fields: CellFields,
  ): FakeCell {
    const rich = toRich(cell);
    if (fields.has("value")) {
      Object.assign(rich, userEnteredToRich(data.userEnteredValue));
    }
    if (fields.has("backgroundColor")) {
      rich.backgroundColor = data.userEnteredFormat?.backgroundColor;
    }
    if (fields.has("numberFormat")) {
      rich.numberFormatType = data.userEnteredFormat?.numberFormat?.type;
    }
    if (fields.has("dataValidation")) {
      rich.dataValidationConditionType = data.dataValidation?.condition?.type;
    }
    return normalized(rich);
  },
  withValue(
    cell: FakeCell,
    value: FakeCellValue,
    isFormula: boolean,
  ): FakeCell {
    return normalized({ ...toRich(cell), value, isFormula });
  },
};

function isCellFieldPath(path: string): path is keyof typeof cellFieldPaths {
  return path in cellFieldPaths;
}

function toRich(cell: FakeCell): FakeRichCellValue {
  if (typeof cell === "object" && cell !== null) return { ...cell };
  return { value: cell };
}

function userEnteredToRich(
  userEntered: ExtendedValue | undefined,
): Pick<FakeRichCellValue, "value" | "isFormula"> {
  if (userEntered?.formulaValue !== undefined) {
    return { value: userEntered.formulaValue, isFormula: true };
  }
  if (userEntered?.errorValue !== undefined) {
    throw new Error("The fake Sheets service does not replay an errorValue.");
  }
  const value =
    userEntered?.stringValue ??
    userEntered?.numberValue ??
    userEntered?.boolValue ??
    null;
  return { value, isFormula: false };
}

// A cell with no fact beyond its value reads back bare, as a fixture writes it.
function normalized(rich: FakeRichCellValue): FakeCell {
  const facts = Object.fromEntries(
    Object.entries(rich).filter(
      ([key, fact]) => key !== "value" && fact !== undefined && fact !== false,
    ),
  );
  if (Object.keys(facts).length === 0) return rich.value;
  return { value: rich.value, ...facts };
}
