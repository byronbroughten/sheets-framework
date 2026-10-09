import { makeTableConfigs } from "../../src/01_SpreadsheetSchema/makeConfigs";

export const tableConfigs = makeTableConfigs({
  "spreadsheetConfig": { "tableId": "spreadsheetConfig", "tableName": "spreadsheetConfig", "sheetGid": 1967106628, "idPrefix": "sscf", "hasIdColumn": false, "hasNameColumn": false },
  "tableConfig": { "tableId": "sheetConfig", "tableName": "tableConfig", "sheetGid": 210603630, "idPrefix": "scf", "hasIdColumn": false, "hasNameColumn": false },
  "columnConfig": { "tableId": "columnConfig", "tableName": "columnConfig", "sheetGid": 2034522667, "idPrefix": "ccf", "hasIdColumn": false, "hasNameColumn": false },
  "valueConfig": { "tableId": "valueConfig", "tableName": "valueConfig", "sheetGid": 2119236084, "idPrefix": "vcf", "hasIdColumn": false, "hasNameColumn": false },
  "item": { "tableId": "item", "tableName": "item", "sheetGid": 1100001, "idPrefix": "itm", "hasIdColumn": true, "hasNameColumn": true },
  "valueTypes": { "tableId": "valueTypes", "tableName": "valueTypes", "sheetGid": 1100002, "idPrefix": "vty", "hasIdColumn": true, "hasNameColumn": false },
  "log": { "tableId": "log", "tableName": "log", "sheetGid": 1100003, "idPrefix": "log", "hasIdColumn": false, "hasNameColumn": false },
  "runItem": { "tableId": "runItem", "tableName": "runItem", "sheetGid": 1100004, "idPrefix": "rit", "hasIdColumn": true, "hasNameColumn": true },
  "computed": { "tableId": "computed", "tableName": "computed", "sheetGid": 1100005, "idPrefix": "cmp", "hasIdColumn": false, "hasNameColumn": false },
  "dates": { "tableId": "dates", "tableName": "dates", "sheetGid": 1100006, "idPrefix": "dat", "hasIdColumn": true, "hasNameColumn": false },
  "layoutLeft": { "tableId": "layoutLeft", "tableName": "layoutLeft", "sheetGid": 1100007, "idPrefix": "lyl", "hasIdColumn": false, "hasNameColumn": false },
  "layoutRight": { "tableId": "layoutRight", "tableName": "layoutRight", "sheetGid": 1100007, "idPrefix": "lyr", "hasIdColumn": false, "hasNameColumn": false }
});
