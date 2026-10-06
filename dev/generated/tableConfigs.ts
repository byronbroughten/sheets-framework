import { makeTableConfigs } from "../../src/01_SpreadsheetSchema/makeConfigs";

export const tableConfigs = makeTableConfigs({
  "spreadsheetConfig": { "tableId": "spreadsheetConfig", "tableName": "spreadsheetConfig", "sheetGid": 1967106628, "idPrefix": "sscf", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": false, "hasNameColumn": false },
  "sheetConfig": { "tableId": "sheetConfig", "tableName": "sheetConfig", "sheetGid": 210603630, "idPrefix": "scf", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": false, "hasNameColumn": false },
  "columnConfig": { "tableId": "columnConfig", "tableName": "columnConfig", "sheetGid": 2034522667, "idPrefix": "ccf", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": false, "hasNameColumn": false },
  "valueConfig": { "tableId": "valueConfig", "tableName": "valueConfig", "sheetGid": 2119236084, "idPrefix": "vcf", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": false, "hasNameColumn": false },
  "item": { "tableId": "item", "tableName": "item", "sheetGid": 1100001, "idPrefix": "itm", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": true, "hasNameColumn": true },
  "valueTypes": { "tableId": "valueTypes", "tableName": "valueTypes", "sheetGid": 1100002, "idPrefix": "vty", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": true, "hasNameColumn": false },
  "log": { "tableId": "log", "tableName": "log", "sheetGid": 1100003, "idPrefix": "log", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": false, "hasNameColumn": false },
  "runItem": { "tableId": "runItem", "tableName": "runItem", "sheetGid": 1100004, "idPrefix": "rit", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": true, "hasNameColumn": true },
  "computed": { "tableId": "computed", "tableName": "computed", "sheetGid": 1100005, "idPrefix": "cmp", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": false, "hasNameColumn": false },
  "dates": { "tableId": "dates", "tableName": "dates", "sheetGid": 1100006, "idPrefix": "dat", "headerRowIndex": 3, "startColIndex": 0, "hasIdColumn": true, "hasNameColumn": false }
});
