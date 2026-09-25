import { makeSheetConfigs } from "../../src/01_SpreadsheetSchema/makeConfigs";

export const sheetConfigs = makeSheetConfigs({
  "spreadsheetConfig": { "sheetGid": 1967106628, "idPrefix": "sscf", "hasIdColumn": false, "hasNameColumn": false },
  "sheetConfig": { "sheetGid": 210603630, "idPrefix": "scf", "hasIdColumn": false, "hasNameColumn": false },
  "columnConfig": { "sheetGid": 2034522667, "idPrefix": "ccf", "hasIdColumn": false, "hasNameColumn": false },
  "valueConfig": { "sheetGid": 2119236084, "idPrefix": "vcf", "hasIdColumn": false, "hasNameColumn": false },
  "item": { "sheetGid": 1100001, "idPrefix": "itm", "hasIdColumn": true, "hasNameColumn": true },
  "valueTypes": { "sheetGid": 1100002, "idPrefix": "vty", "hasIdColumn": true, "hasNameColumn": false },
  "log": { "sheetGid": 1100003, "idPrefix": "log", "hasIdColumn": false, "hasNameColumn": false },
  "runItem": { "sheetGid": 1100004, "idPrefix": "rit", "hasIdColumn": true, "hasNameColumn": true },
  "computed": { "sheetGid": 1100005, "idPrefix": "cmp", "hasIdColumn": false, "hasNameColumn": false },
  "dates": { "sheetGid": 1100006, "idPrefix": "dat", "hasIdColumn": true, "hasNameColumn": false }
});
