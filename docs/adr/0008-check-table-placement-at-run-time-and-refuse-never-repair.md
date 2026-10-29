# Check a Table's placement at run time, and refuse a misplaced one rather than repair it

A Table's range lives in a spreadsheet a person edits by hand, so no type can keep it where the layout requires, and a runtime check of each managed Table after each fetch is the only instrument there, not a lapse from structural guarantees. A misplaced Table stops the run with a message giving where it is and where it must be, kept distinct from an absent one because the fixes are opposite: a moved Table had read as absent, with advice to insert a Table that would have the operator rebuild a working one. On the already-aborting path one extra properties read tells the two apart. Moving the Table back was available and rejected, because a moved Table usually means the operator restructured the sheet on purpose.

Source: byronbroughten/byro-repo#9, byronbroughten/sheets-framework#62
