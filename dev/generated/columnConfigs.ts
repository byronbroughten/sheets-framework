import { makeColumnConfigs } from "@byronbroughten/sheets-framework";

export const columnConfigs = makeColumnConfigs({
  "valueConfig": {
    "exampleValue": { "columnId": "c:vcf:AYSrgZf", "header": "Example value", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null }
  },
  "item": {
    "optionalNote": { "columnId": "c:itm:optionalNote", "header": "Optional note", "valueName": "string", "isFormula": false, "emptyValueAllowed": true, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "requiredCount": { "columnId": "c:itm:requiredCount", "header": "Required count", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DOUBLE","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "id": { "columnId": "c:itm:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": {"kind":"checkbox"}, "groupHeading1": null },
    "name": { "columnId": "c:itm:name", "header": "Name", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null }
  },
  "valueTypes": {
    "id": { "columnId": "c:vty:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": {"kind":"checkbox"}, "groupHeading1": null },
    "stringValue": { "columnId": "c:vty:stringValue", "header": "String value", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "numberValue": { "columnId": "c:vty:numberValue", "header": "Number value", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DOUBLE","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "dateValue": { "columnId": "c:vty:dateValue", "header": "Date value", "valueName": "date", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DATE","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "sampledBoolean": { "columnId": "c:vty:sampledBoolean", "header": "Sampled boolean", "valueName": "boolean", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":null,"rule":null}, "actionRowCell": null, "groupHeading1": null },
    "checkbox": { "columnId": "c:vty:checkbox", "header": "Checkbox", "valueName": "checkbox", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"BOOLEAN","rule":null}, "actionRowCell": null, "groupHeading1": null }
  },
  "log": {
    "entry": { "columnId": "c:log:entry", "header": "Entry", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "amount": { "columnId": "c:log:amount", "header": "Amount", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DOUBLE","rule":null}, "actionRowCell": null, "groupHeading1": null }
  },
  "runItem": {
    "id": { "columnId": "c:rit:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": {"kind":"checkbox"}, "groupHeading1": null },
    "name": { "columnId": "c:rit:name", "header": "Name", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "selected": { "columnId": "c:rit:selected", "header": "Selected", "valueName": "checkbox", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"BOOLEAN","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "result": { "columnId": "c:rit:result", "header": "Result", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": {"kind":"checkbox"}, "groupHeading1": null },
    "startTime": { "columnId": "c:rit:startTime", "header": "Start time", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "runStatus": { "columnId": "c:rit:runStatus", "header": "Run status", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null }
  },
  "computed": {
    "amount": { "columnId": "c:cmp:amount", "header": "Amount", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DOUBLE","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "rowNumber": { "columnId": "c:cmp:rowNumber", "header": "Row number", "valueName": "number", "isFormula": true, "emptyValueAllowed": false, "customDefaultValue": null, "formula": "=ROW()", "columnType": {"type":"DOUBLE","rule":null}, "actionRowCell": null, "groupHeading1": null }
  },
  "dates": {
    "requiredDate": { "columnId": "c:dat:requiredDate", "header": "Required date", "valueName": "date", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DATE","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "optionalDate": { "columnId": "c:dat:optionalDate", "header": "Optional date", "valueName": "date", "isFormula": false, "emptyValueAllowed": true, "customDefaultValue": null, "formula": null, "columnType": {"type":"DATE","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "id": { "columnId": "c:dat:id", "header": "ID", "valueName": "id", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": {"kind":"checkbox"}, "groupHeading1": null }
  },
  "layoutLeft": {
    "entry": { "columnId": "c:lyl:entry", "header": "Entry", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "amount": { "columnId": "c:lyl:amount", "header": "Amount", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DOUBLE","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "note": { "columnId": "c:lyl:note", "header": "Note", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":null,"rule":null}, "actionRowCell": null, "groupHeading1": null }
  },
  "layoutRight": {
    "entry": { "columnId": "c:lyr:entry", "header": "Entry", "valueName": "string", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"TEXT","rule":null}, "actionRowCell": null, "groupHeading1": null },
    "amount": { "columnId": "c:lyr:amount", "header": "Amount", "valueName": "number", "isFormula": false, "emptyValueAllowed": false, "customDefaultValue": null, "formula": null, "columnType": {"type":"DOUBLE","rule":null}, "actionRowCell": null, "groupHeading1": null }
  }
});
