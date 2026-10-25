# Assign each sheet's ID prefix once, and sample it from its column IDs

A typed ID prefix was a second copy of what every column ID on the sheet already carries, so Sheet Config lost its ID prefix column and the "ID prefix is unique or empty" helper that checked it. Config sync now generates a readable prefix from the tab title when it first mints a sheet's column IDs, and from then on samples it from the column ID row. Hidden sheet metadata was rejected as a new kind of storage the probe can't show, and reading the prefix from the last generated configs was rejected because the cache would become an input to its own regeneration, losing prefixes on a reset and letting two branches assign the same one. Mixed prefixes on one sheet, or one prefix on two sheets, stop the run rather than being repaired.

Source: byronbroughten/byro-repo#87
