export interface IdNameMapTable {
  name: string;
  sheetGid: number;
  columns?: Record<string, string>;
}
export interface IdNameMapTab {
  // Optional only until each package's map is regenerated with it.
  name?: string;
  title: string;
  tableIds: string[];
}
// Tables are keyed by Table ID and tabs by GID, the two keys a live spreadsheet keeps unique.
export interface IdNameMapGeneric {
  tables: Record<string, IdNameMapTable>;
  tabs: Record<string, IdNameMapTab>;
}
