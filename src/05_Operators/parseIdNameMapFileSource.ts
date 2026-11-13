import { Val } from "@byronbroughten/utils/val";

import type {
  IdNameMapGeneric,
  IdNameMapTab,
  IdNameMapTable,
} from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { idNameMapOpening } from "./configFileSource";

// The map is a type with no runtime value, so sync reads its previous one back from the source text.
export function parseIdNameMapFileSource(source: string): IdNameMapGeneric {
  const openingIndex = source.indexOf(idNameMapOpening);
  if (openingIndex === -1) {
    throw new Error("The ID-name map source has no IdNameMap interface.");
  }
  const parsed: unknown = JSON.parse(
    source.slice(openingIndex + idNameMapOpening.length),
  );
  const idNameMap = validateRecord(parsed, "ID-name map");
  return {
    tables: mapRecord(
      validateRecord(idNameMap.tables, "ID-name map tables"),
      validateTable,
    ),
    tabs: mapRecord(
      validateRecord(idNameMap.tabs, "ID-name map tabs"),
      validateTab,
    ),
  };
}

function validateTable(value: unknown): IdNameMapTable {
  const table = validateRecord(value, "ID-name map Table");
  const name = Val.validate.string(table.name);
  const sheetGid = Val.validate.number(table.sheetGid);
  if (table.columns === undefined) return { name, sheetGid };
  return {
    name,
    sheetGid,
    columns: mapRecord(
      validateRecord(table.columns, "ID-name map columns"),
      Val.validate.string,
    ),
  };
}

function validateTab(value: unknown): IdNameMapTab {
  const tab = validateRecord(value, "ID-name map tab");
  if (!Array.isArray(tab.tableIds)) {
    throw new Error("An ID-name map tab's tableIds is not an array.");
  }
  const tableIds = tab.tableIds.map((tableId: unknown) =>
    Val.validate.string(tableId),
  );
  return {
    name: Val.validate.string(tab.name),
    title: Val.validate.string(tab.title),
    tableIds,
  };
}

function validateRecord(
  value: unknown,
  label: string,
): Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new Error(`The ${label} is not an object.`);
  }
  return Object.fromEntries(Object.entries(value));
}

function mapRecord<VL>(
  record: Record<string, unknown>,
  validate: (value: unknown) => VL,
): Record<string, VL> {
  return Object.entries(record).reduce<Record<string, VL>>(
    (mapped, [key, value]) => {
      mapped[key] = validate(value);
      return mapped;
    },
    {},
  );
}
