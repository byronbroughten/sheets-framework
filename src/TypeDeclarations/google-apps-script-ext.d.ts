declare namespace GoogleAppsScript {
  namespace Sheets {
    namespace Collection {
      interface SpreadsheetsCollection {
        getByDataFilter(
          resource: Schema.GetSpreadsheetByDataFilterRequest,
          spreadsheetId: string,
          optionalArgs?: { fields?: string },
        ): Schema.Spreadsheet;
      }
    }
    namespace Schema {
      // -----------------------------------------------------------------
      // 1. Table Data Structures
      // -----------------------------------------------------------------

      export interface TableColumnProperties {
        columnIndex?: number;
        columnName?: string;
        columnType?: string;
        dataValidationRule?: DataValidationRule;
      }

      export interface Table {
        tableId?: string;
        name?: string;
        range?: GoogleGridRange;
        columnProperties?: TableColumnProperties[];
      }

      // -----------------------------------------------------------------
      // 2. Request Objects
      // -----------------------------------------------------------------

      export interface AppendCellsRequest {
        sheetId?: number;
        tableId?: string;
        rows?: RowData[];
        fields?: string;
      }

      export interface UpdateTableRequest {
        table?: Table;
        fields?: string;
      }

      export interface AddTableRequest {
        table?: Table;
      }

      // Merge into the global Schema.Request interface
      export interface Request {
        addTable?: AddTableRequest;
        appendCells?: AppendCellsRequest;
        updateTable?: UpdateTableRequest;
      }

      // -----------------------------------------------------------------
      // 3. Extending SheetNamed Interface
      // -----------------------------------------------------------------

      export interface Sheet {
        tables?: Table[];
      }
    }

    // -------------------------------------------------------------------
    // 4. Constructor Helper Functions (Optional but recommended)
    // -------------------------------------------------------------------

    export interface SheetsCollection {
      newAppendCellsRequest?: () => Schema.AppendCellsRequest;
      newUpdateTableRequest?: () => Schema.UpdateTableRequest;
      newTable?: () => Schema.Table;
      newTableColumnProperties?: () => Schema.TableColumnProperties;
    }
  }
}
