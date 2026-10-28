# Seeing the raw Sheets JSON

`GoogleSheetsAPI` maps the payload before anything can log it, and the `gworkspace` MCP returns cell values only. To see what Google actually sent, use the committed probe:

```
sheets-framework probe --fields 'sheets(properties(sheetId,title),protectedRanges)'
sheets-framework probe --filter '{"dataFilters":[...]}' [--fields '<mask>']
sheets-framework probe --path sheets.title=Log.protectedRanges      # re-read the last response, no request
```

The bin's `probe` (`scripts/sheetsProbe.ts`) sends one request through the Node host's `SheetsTransport`, authenticated like a chore. That request is a `GET` with a `fields` mask or a `:getByDataFilter`, and the script cannot build any other kind. It writes the full response, pretty-printed, to the gitignored `.probe/last.json` in its package: `.probe/last.json` for the app, `dev/.probe/last.json` for dev. Stdout gets only a summary: top-level keys, array counts, and each sheet's id and title. With `--path`, stdout gets that one subtree, printed whole when it is short and summarized when it is not. When the summary isn't enough, `Read` a line range of the file it names. The throwaway `scripts/*.tmp.mjs` route this replaced is retired.
