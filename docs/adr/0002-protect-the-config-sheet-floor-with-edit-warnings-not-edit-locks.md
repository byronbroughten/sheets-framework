---
status: superseded by sheets-framework#219, which removes the config-sheet floor
---

# Protect the config-sheet floor with edit warnings, not edit locks

A lock naming no editors inherits every account the spreadsheet is shared with and stops none of them, and the only lock that keeps another editor out names editors explicitly, which needs an address list with nowhere to live. So the floor carries edit warnings and no edit locks, and the framework keeps its lock capability for other callers. Locks on the floor wait for an Admin Config to hold those addresses, and for an answer to why a probe lock naming the owner blocked the owner.

## Superseded

[sheets-framework#219](https://github.com/byronbroughten/sheets-framework/issues/219) removes the three config tabs and with them the floor these warnings protected. Intended state holds no protections at all, for the reasons in the deliberate-absence page's protections bullet ([deliberate absence](../design/deliberate-absence.md)).

Source: byronbroughten/byro-repo#49, byronbroughten/byro-repo#51
