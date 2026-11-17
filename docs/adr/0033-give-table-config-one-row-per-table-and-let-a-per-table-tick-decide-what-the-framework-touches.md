---
status: superseded by sheets-framework#219, which removes Table Config
---

# Give Table Config one row per Table, and let a per-Table tick decide what the framework touches

Table Config replaced Sheet Config with one row per Google Table, identified by its Table ID, and Let api access moved from the sheet to the Table as that row's only declared column; runs and `gen:configs` touch only ticked Tables. Access moved to the Table so that an operator's own Tables on a managed sheet stay theirs, left alone by runs and by config generation.

## Superseded

[sheets-framework#219](https://github.com/byronbroughten/sheets-framework/issues/219) removes Table Config and its tick. API access stays per Table, so an operator's own Tables stay theirs, but it is now a Table's presence in the generated `tableConfigs`, changed only by a `letApiAccess` alteration ([granting and revoking](../generated-data/table-configs.md#granting-and-revoking-api-access)).

Source: byronbroughten/sheets-framework#108
