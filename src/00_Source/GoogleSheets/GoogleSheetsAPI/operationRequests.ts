import type {
  LocalWriteOperation,
  OpaqueRawRequest,
  TableColumnPropertiesAdd,
  TableColumnPropertiesUpdate,
  UpdateTabPropertiesOperation,
} from "../../RawSource/RawSource.ts";
import type { GoogleRequest, ModeledRequest } from "../GoogleSheetsAPI.ts";
import { cellDataRequests } from "./cellData.ts";
import { googleConditionalFormatRule } from "./conditionalFormats.ts";
import { googleColor } from "./googleColor.ts";
import { googleProtectedRange } from "./protectedRanges.ts";

type GoogleColumnProperties =
  GoogleAppsScript.Sheets.Schema.TableColumnProperties;

function unwrapRawRequest(request: OpaqueRawRequest): GoogleRequest {
  return request as unknown as GoogleRequest;
}

export function localOperationToGoogleRequests(
  operation: LocalWriteOperation,
): GoogleRequest[] {
  if (operation.kind === "raw") return [unwrapRawRequest(operation.request)];
  return modeledOperationToGoogleRequests(operation);
}

function modeledOperationToGoogleRequests(
  operation: Exclude<LocalWriteOperation, { kind: "raw" }>,
): ModeledRequest[] {
  switch (operation.kind) {
    case "appendDimension":
      return [
        {
          appendDimension: {
            sheetId: operation.sheetId,
            dimension: operation.dimension,
            length: operation.addedCount,
          },
        },
      ];
    case "insertRange":
      return [
        {
          insertRange: {
            range: operation.range,
            shiftDimension: operation.shiftDimension,
          },
        },
      ];
    case "updateTableRange":
      return [
        {
          updateTable: {
            table: { tableId: operation.tableId, range: operation.range },
            fields: "range",
          },
        },
      ];
    case "deleteRange":
      return [
        {
          deleteRange: {
            range: operation.range,
            shiftDimension: operation.shiftDimension,
          },
        },
      ];
    case "copyPaste":
      return [
        {
          copyPaste: {
            source: operation.source,
            destination: operation.destination,
            pasteType: operation.pasteType,
            pasteOrientation: "NORMAL",
          },
        },
      ];
    case "cutPaste":
      return [
        {
          cutPaste: {
            source: operation.source,
            destination: operation.destination,
            pasteType: "PASTE_NORMAL",
          },
        },
      ];
    case "fillColumn":
      return cellDataRequests.fillColumn(operation);
    case "fillCell":
      return cellDataRequests.fillCell(operation);
    case "findReplace":
      return [
        {
          findReplace: { ...operation.terms, ...operation.scope },
        },
      ];
    case "deleteTableRows":
      return [
        {
          deleteRange: { range: operation.range, shiftDimension: "ROWS" },
        },
      ];
    case "sortTable":
      return [
        {
          sortRange: {
            range: operation.range,
            sortSpecs: [
              {
                dimensionIndex: operation.colIdxToSortBy,
                sortOrder: operation.sortOrder,
              },
            ],
          },
        },
      ];
    case "addConditionalFormatRule":
      return [
        {
          addConditionalFormatRule: {
            index: operation.index,
            rule: googleConditionalFormatRule.fromModelable(operation.rule),
          },
        },
      ];
    case "updateConditionalFormatRule":
      return [
        {
          updateConditionalFormatRule: {
            sheetId: operation.sheetId,
            index: operation.index,
            rule: googleConditionalFormatRule.fromModelable(operation.rule),
          },
        },
      ];
    case "deleteConditionalFormatRule":
      return [
        {
          deleteConditionalFormatRule: {
            sheetId: operation.sheetId,
            index: operation.index,
          },
        },
      ];
    case "addProtectedRange":
      return [
        {
          addProtectedRange: {
            protectedRange: googleProtectedRange.fromContent(
              operation.protection,
            ),
          },
        },
      ];
    case "deleteProtectedRange":
      return [
        {
          deleteProtectedRange: {
            protectedRangeId: operation.protectedRangeId,
          },
        },
      ];
    case "addSheet":
      return [
        {
          addSheet: {
            properties: {
              sheetId: operation.sheetId,
              title: operation.title,
              gridProperties: {
                rowCount: operation.rowCount,
                columnCount: operation.columnCount,
              },
            },
          },
        },
      ];
    case "deleteSheet":
      return [{ deleteSheet: { sheetId: operation.sheetId } }];
    // Google writes addTable's own column names three columns right of the Table; observed: docs/sheets-behaviour.md#creating-and-updating-tables.
    case "addTable":
      return [
        {
          addTable: {
            table: {
              tableId: operation.tableId,
              name: operation.name,
              range: operation.range,
            },
          },
        },
        columnPropertiesRequest(operation.tableId, operation.columnProperties),
      ];
    // deleteTable leaves the head rows, so an updateCells with no rows clears their values and validation; observed: docs/sheets-behaviour.md#creating-and-updating-tables.
    case "deleteTable":
      return [
        { deleteTable: { tableId: operation.tableId } },
        {
          updateCells: {
            range: operation.headRows,
            fields: "userEnteredValue,dataValidation",
          },
        },
      ];
    case "renameSheet":
      return [
        {
          updateSheetProperties: {
            properties: {
              sheetId: operation.sheetId,
              title: operation.title,
            },
            fields: "title",
          },
        },
      ];
    case "updateTabProperties":
      return [tabPropertiesRequest(operation)];
    case "renameTable":
      return [
        {
          updateTable: {
            table: {
              tableId: operation.tableId,
              name: operation.name,
            },
            fields: "name",
          },
        },
      ];
    case "setTableColumnProperties":
      return [
        columnPropertiesRequest(operation.tableId, operation.columnProperties),
      ];
    case "addCheckboxValidation":
      return [
        {
          setDataValidation: {
            range: operation.range,
            rule: { condition: { type: "BOOLEAN" } },
          },
        },
      ];
    default: {
      const exhaustive: never = operation;
      throw new Error(
        `Unknown local write operation: ${JSON.stringify(exhaustive)}`,
      );
    }
  }
}

type TabProperty = keyof UpdateTabPropertiesOperation["properties"];

// The field each tab property sets; the mask names only what the operation carries, so an absent property is left alone.
const tabPropertyMasks: [TabProperty, string][] = [
  ["color", "tabColorStyle"],
  ["hidden", "hidden"],
  ["frozenRows", "gridProperties.frozenRowCount"],
  ["frozenColumns", "gridProperties.frozenColumnCount"],
  ["index", "index"],
];

function tabPropertiesRequest({
  sheetId,
  properties,
}: UpdateTabPropertiesOperation): ModeledRequest {
  const { color, hidden, frozenRows, frozenColumns, index } = properties;
  const fields = tabPropertyMasks.flatMap(([property, mask]) =>
    properties[property] === undefined ? [] : [mask],
  );
  return {
    updateSheetProperties: {
      properties: {
        sheetId,
        tabColorStyle:
          color === undefined ? undefined : googleColor.fromTabColor(color),
        hidden,
        gridProperties: {
          frozenRowCount: frozenRows,
          frozenColumnCount: frozenColumns,
        },
        index,
      },
      fields: fields.join(","),
    },
  };
}

function columnPropertiesRequest(
  tableId: string,
  columnProperties: TableColumnPropertiesAdd[] | TableColumnPropertiesUpdate[],
): ModeledRequest {
  return {
    updateTable: {
      table: { tableId, columnProperties: columnProperties.map(googleColumn) },
      fields: "columnProperties",
    },
  };
}

// A rule with no values sends none, as the Table fetch reads one back.
function googleColumn(
  column: TableColumnPropertiesAdd,
): GoogleColumnProperties {
  const { dataValidationRule, ...rest } = column;
  if (dataValidationRule === undefined) return rest;
  const { condition, values } = dataValidationRule;
  return {
    ...rest,
    dataValidationRule: {
      condition: {
        type: condition,
        ...(values.length > 0
          ? { values: values.map((userEnteredValue) => ({ userEnteredValue })) }
          : {}),
      },
    },
  };
}
