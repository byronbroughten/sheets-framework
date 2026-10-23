# Keep Google Sheets inside one platform module

`src/00_Source/GoogleSheets/` is the one place that knows the spreadsheet product is Google Sheets, so an Excel (Office Scripts) implementation can replace that folder rather than chase Google types through the tiers. A platform capability reaches neutral code as a neutral type such as `SheetEdit`, the adapter is bound to its spreadsheet when it is built, and a lint rule, not a convention, fences the module.

Source: byronbroughten/byro-repo#85, byronbroughten/byro-repo#104
