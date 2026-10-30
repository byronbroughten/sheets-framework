# Hold the Raw tier's shared Table members between its base class and its concrete classes

The members both concrete Raw Table classes call live on an abstract class, now `TableCommonRaw`, between the tier's base class and those classes, rather than on the base class. The row, column and cell base classes hang off that base, so two of the shared members would be illegal overrides of same-named members on the row and cell classes, and the rest would be inherited by classes with no use for them; the extra class gives Raw the shape the other tiers already had.

Source: byronbroughten/byro-repo#6
