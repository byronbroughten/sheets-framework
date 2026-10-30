# Read the sheet through the Sheets API, not `SpreadsheetApp`, even in a trigger

A trigger already has the sheet open, so `SpreadsheetApp` looked like the cheaper read, but each of its calls is its own round trip, and the calls that would replace one gathered `getByDataFilter` measured slower than it. So reads stay on the Sheets API, gathered into one request, and `SpreadsheetApp` isn't re-proposed without new measurements; the numbers are in `docs/architecture/round-trips.md`.

Source: commit byronbroughten/byro-repo@933351eb
