import { makeColumnConfigs } from "../../src/01_SpreadsheetSchema/makeConfigs";

export const columnConfigs = makeColumnConfigs({
  "spreadsheetConfig": {
    "tableMenuSpace": { "columnId": "c:sscf:0xzKfv_", "header": "Table menu space", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "fillRowIdsTimeLastRan": { "columnId": "c:sscf:eV73Th5", "header": "Fill row IDs, time last ran", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "fillRowIdsRunStatus": { "columnId": "c:sscf:ebB4-9S", "header": "Fill row IDs, run status", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "syncConfigSheetRowsTimeLastRan": { "columnId": "c:sscf:pLBSdae", "header": "Sync config sheet rows, time last ran", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "syncConfigSheetRowsRunStatus": { "columnId": "c:sscf:W-JABu_", "header": "Sync config sheet rows, run status", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "sheetConfig": {
    "sheetGid": { "columnId": "c:scf:WgnoW8d", "header": "Sheet GID", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "sheetTitle": { "columnId": "c:scf:0Ctj9xZ", "header": "Sheet title", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "letApiAccess": { "columnId": "c:scf:GOJ0ixi", "header": "Let api access", "valueName": "checkbox", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "columnConfig": {
    "sheetGid": { "columnId": "c:ccf:1-6AQIj", "header": "Sheet GID", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "columnId": { "columnId": "c:ccf:vj9_rre", "header": "Column ID", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "sheetTitle": { "columnId": "c:ccf:949GjdB", "header": "Sheet title", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "header": { "columnId": "c:ccf:kqA31oK", "header": "Header", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "emptyValueAllowed": { "columnId": "c:ccf:volkLl6", "header": "Empty value allowed", "valueName": "checkbox", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "valueConfig": {
    "exampleValue": { "columnId": "c:vcf:AYSrgZf", "header": "Example value", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "item": {
    "optionalNote": { "columnId": "c:itm:optionalNote", "header": "Optional note", "valueName": "string", "isFormula": false, "emptyValueAllowed": true, "customDefaultValue": null },
    "requiredCount": { "columnId": "c:itm:requiredCount", "header": "Required count", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "id": { "columnId": "c:itm:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "name": { "columnId": "c:itm:name", "header": "Name", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "valueTypes": {
    "id": { "columnId": "c:vty:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "stringValue": { "columnId": "c:vty:stringValue", "header": "String value", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "numberValue": { "columnId": "c:vty:numberValue", "header": "Number value", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "dateValue": { "columnId": "c:vty:dateValue", "header": "Date value", "valueName": "date", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "sampledBoolean": { "columnId": "c:vty:sampledBoolean", "header": "Sampled boolean", "valueName": "boolean", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "checkbox": { "columnId": "c:vty:checkbox", "header": "Checkbox", "valueName": "checkbox", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "log": {
    "entry": { "columnId": "c:log:entry", "header": "Entry", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "amount": { "columnId": "c:log:amount", "header": "Amount", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "runItem": {
    "id": { "columnId": "c:rit:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "name": { "columnId": "c:rit:name", "header": "Name", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "selected": { "columnId": "c:rit:selected", "header": "Selected", "valueName": "checkbox", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "result": { "columnId": "c:rit:result", "header": "Result", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "startTime": { "columnId": "c:rit:startTime", "header": "Start time", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "runStatus": { "columnId": "c:rit:runStatus", "header": "Run status", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "computed": {
    "amount": { "columnId": "c:cmp:amount", "header": "Amount", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "rowNumber": { "columnId": "c:cmp:rowNumber", "header": "Row number", "valueName": "number", "isFormula": true, "emptyValueAllowed": false, "customDefaultValue": null }
  },
  "dates": {
    "requiredDate": { "columnId": "c:dat:requiredDate", "header": "Required date", "valueName": "date", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null },
    "optionalDate": { "columnId": "c:dat:optionalDate", "header": "Optional date", "valueName": "date", "isFormula": false, "emptyValueAllowed": true, "customDefaultValue": null },
    "id": { "columnId": "c:dat:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null }
  }
});
