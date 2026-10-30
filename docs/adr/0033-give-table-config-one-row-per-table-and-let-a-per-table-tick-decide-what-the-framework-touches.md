# Give Table Config one row per Table, and let a per-Table tick decide what the framework touches

Table Config replaced Sheet Config with one row per Google Table in the spreadsheet, `Table ID | Table name | Sheet title | Let api access`: Table ID is the row's identity, Table name and Sheet title are readable identity the correction pass rewrites, and Let api access is the only declared column. The sync appends a new Table unticked and gives a tab with no Table no row, and `gen:configs` samples only ticked Tables, so an unticked Table gets no Column Config rows and no `tableConfigs` entry. Access moved from the sheet to the Table so that an operator's own Tables on a managed sheet stay theirs, left alone by runs and by config generation.

Source: byronbroughten/sheets-framework#108
