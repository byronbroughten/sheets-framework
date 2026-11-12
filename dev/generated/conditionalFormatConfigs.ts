import { makeConditionalFormatConfigs } from "@byronbroughten/sheets-framework";

export const conditionalFormatConfigs = makeConditionalFormatConfigs({
  "tables": {},
  "tabs": {
    "scratch217A": {
      "C5:C7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217001, "ranges": ["C5:C7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:F7 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217001, "ranges": ["A5:F7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "A4:A7 | =LEN(A4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217001, "ranges": ["A4:A7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(A4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E7 | =C5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217001, "ranges": ["E5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=C5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217B": {
      "C5:C7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217002, "ranges": ["C5:C7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:E7 | =$B5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217002, "ranges": ["A5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$B5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "A4:A7 | =LEN(A4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217002, "ranges": ["A4:A7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(A4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E7 | =C5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217002, "ranges": ["E5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=C5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217C": {
      "A5:A7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217003, "ranges": ["A5:A7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:F7 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217003, "ranges": ["A5:F7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "C4:C7 | =LEN(C4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217003, "ranges": ["C4:C7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(C4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E7 | =A5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217003, "ranges": ["E5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=A5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217D": {
      "A5:A7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217004, "ranges": ["A5:A7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:E7 | =$D5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217004, "ranges": ["A5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$D5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "C4:C7 | =LEN(C4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217004, "ranges": ["C4:C7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(C4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E7 | =A5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217004, "ranges": ["E5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=A5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217E": {
      "E5:E7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217005, "ranges": ["E5:E7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:F7 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217005, "ranges": ["A5:F7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "B4:B7 | =LEN(B4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217005, "ranges": ["B4:B7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(B4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "F5:F7 | =E5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217005, "ranges": ["F5:F7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=E5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217F": {
      "A5:E7 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217006, "ranges": ["A5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "B4:B7 | =LEN(B4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217006, "ranges": ["B4:B7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(B4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "D5:D7 | =D5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217006, "ranges": ["D5:D7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=D5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217G": {
      "C5:C7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217007, "ranges": ["C5:C7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:E7 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217007, "ranges": ["A5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "B4:B7 | =LEN(B4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217007, "ranges": ["B4:B7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(B4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "D5:D7 | =C5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217007, "ranges": ["D5:D7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=C5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217H": {
      "D5:D7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217008, "ranges": ["D5:D7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:E7 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217008, "ranges": ["A5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "B4:B7 | =LEN(B4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217008, "ranges": ["B4:B7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(B4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E7 | =D5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217008, "ranges": ["E5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=D5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217J": {
      "D5:D8 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217010, "ranges": ["D5:D8"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:E8 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217010, "ranges": ["A5:E8"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "B4:B8 | =LEN(B4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217010, "ranges": ["B4:B8"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(B4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E8 | =D5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217010, "ranges": ["E5:E8"], "condition": {"type":"CUSTOM_FORMULA","formula":"=D5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217K": {
      "C5:C7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217011, "ranges": ["C5:C7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:E7 | =$B5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217011, "ranges": ["A5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$B5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "A4:A7 | =LEN(A4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217011, "ranges": ["A4:A7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(A4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E7 | =C5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217011, "ranges": ["E5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=C5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    },
    "scratch217L": {
      "A5:A7 | NUMBER_GREATER 2 | background #ffcccc": { "kind": "boolean", "sheetGid": 217012, "ranges": ["A5:A7"], "condition": {"type":"NUMBER_GREATER","value":2}, "format": {"backgroundColor":"#ffcccc"} },
      "A5:H7 | =$C5=\"x\" | background #ccffcc": { "kind": "boolean", "sheetGid": 217012, "ranges": ["A5:H7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$C5=\"x\""}, "format": {"backgroundColor":"#ccffcc"} },
      "C4:C7 | =LEN(C4)>5 | background #ccccff": { "kind": "boolean", "sheetGid": 217012, "ranges": ["C4:C7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=LEN(C4)>5"}, "format": {"backgroundColor":"#ccccff"} },
      "E5:E7 | =B5>2 | background #ffffb2": { "kind": "boolean", "sheetGid": 217012, "ranges": ["E5:E7"], "condition": {"type":"CUSTOM_FORMULA","formula":"=B5>2"}, "format": {"backgroundColor":"#ffffb2"} }
    }
  }
});
