# Repair a payload's omitted rows in Raw's finalize pass, narrowed to the Table

Sheets omits `rowData` for a fetched row that holds no value, formula or number format, so a sheet left in its designed blank-row state had no active facts and crashed the config sync. The finalize pass that already backfills omitted cells now also ensures blank facts for every Table column on a fetched top data row or a fetched full column, never overwriting one the payload supplied, so fetched-and-empty reads as an empty cell in a populated row does and a never-fetched column still throws, naming its sheet. Moving every fact into finalize was rejected because formula-ness and number format exist only while the payload is in hand, and both writers cover the Table's columns only, since the payload describes every grid column (23 against a 12-column Table on the reported sheet).

Source: byronbroughten/byro-repo#17
