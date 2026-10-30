# Identify a Table by its live `tableId`, and mint a random one for a Table the framework creates

A Table is identified by its live `tableId`. Its config key, the generic behind `ss.table(name)`, is derived from its live name, and the name is stored for display and for `reference` and `single` only, never as identity, because live names already drift from config keys and a rename rewrites every reference to it. A column is identified by its column ID alone, which its prefix makes unique across the spreadsheet. Two Tables whose names give the same key fail `gen:configs`, naming both, with no automatic suffix. A Table the framework creates gets a random `tableId` the framework chooses, rather than its name, which had been sent as its ID, or one Google mints, which arrives only in the reply, too late for the `updateTable` that names its columns in the same batch.

Source: byronbroughten/sheets-framework#51, byronbroughten/sheets-framework#59
