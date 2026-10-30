# Enforce a chore's dry run at the Node host's adapter

In dry-run mode the Node host's adapter, the single door to Google, records each write request and returns an empty response, so no flush, the config orchestrator's internal one included, can reach the sheet, and a dry run that writes is unrepresentable rather than discouraged. Withholding the flush in the chore runner would miss framework code that flushes on its own account, and narrowing a chore's spreadsheet handle to remove the flush was dropped as buying nothing the adapter doesn't already give. Reads are not gated, so a dry run fetches from the live sheet.

Source: byronbroughten/byro-repo#25
