# Give Table Config one row per Table, and let a per-Table tick decide what the framework touches

Table Config replaced Sheet Config with one row per Google Table, identified by its Table ID, and Let api access moved from the sheet to the Table as that row's only declared column; runs and `gen:configs` touch only ticked Tables ([the sheet and its sync](../generated-data/table-configs.md)). Access moved to the Table so that an operator's own Tables on a managed sheet stay theirs, left alone by runs and by config generation.

Source: byronbroughten/sheets-framework#108
