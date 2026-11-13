# Hold entered-data loss at the Node host's adapter until a hash-bound confirm

Before each flush, the Node host's adapter classifies every request, and a flush that would delete or overwrite **entered data**, or carries another held item (a forced property, a dropdown style reset under `"ask"`, `findReplace`, an unmeasured effect, an unclassified verb), is sent only when `--confirm <token>` covers it; the token is a hash of the target and the held list a dry run printed. Only the dry run prints the held list and the token; a refused send says only "rerun the dry run", so the token a send needs always comes from a dry run against the sheet as it is. So the framework promises that only a command holding `--confirm` can lose entered data, which a harness can gate on from the command text without knowing how the classification works, and the framework knows nothing about any agent or gate.

## What it reads and what it skips

Each cell is judged from a read the run already makes (`formulaValue` added to Node-host reads; `sync` reads only the columns a delete or formula alteration touches), and a cell no read covered is held. Endpoints, which a person triggers on the Apps Script host and whose deletes are their design, are not classified.

## Considered options

- **Replaying the requests on the fake spreadsheet and diffing**: rejected, since it costs a seeded copy of the live sheet per send and fails open wherever the fake's fidelity falls short.
- **A column-wide formula trait from the generated configs, set from every data cell or from a sample of a few**: rejected. Reading every cell grows with the data; a sample frees every overwrite in a column it calls formula, and the cell it misses is a person's hand override of one computed value, a loss nobody sees.
- **A bare `--allow-data-loss` flag**: rejected, since it approves whatever the sheet holds when the command runs, not the list the developer saw.
- **Classifying a whole dry run up front, then sending**: rejected, since a later flush is planned from reads that an earlier flush's writes would have changed; flush by flush, a refusal at a later flush leaves the earlier ones landed and reported.

Source: byronbroughten/sheets-framework#188, byronbroughten/sheets-framework#178
