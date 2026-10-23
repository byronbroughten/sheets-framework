# Protect the config-sheet floor with edit warnings, not edit locks

A lock naming no editors inherits every account the spreadsheet is shared with and stops none of them, and the only lock that keeps another editor out names editors explicitly, which needs an address list with nowhere to live. So the floor carries edit warnings and no edit locks, and the framework keeps its lock capability for other callers. Locks on the floor wait for an Admin Config to hold those addresses, and for an answer to why a probe lock naming the owner blocked the owner.

Source: byronbroughten/byro-repo#49, byronbroughten/byro-repo#51
