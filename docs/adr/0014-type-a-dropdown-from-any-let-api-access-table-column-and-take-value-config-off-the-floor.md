# Type a dropdown from any Let api access Table column, and take Value Config off the floor

Any column of a Let api access Table can be a value source for a dropdown whose range points at it, named for the source Table plus its header. A Name or ID column source is a live reference typed `string`, because its options are rows that grow without a regen, so a name with no matching row is the endpoint's to report and the framework adds no read-time check. Every other source is a fixed list frozen at the config sync as a literal union. With lists free to live on any Table, Value Config left the config-sheet floor: the framework no longer creates, restores or seeds it, settling for now the floor question that byronbroughten/byro-repo#49 and byronbroughten/byro-repo#99 went back and forth on.

Source: byronbroughten/sheets-framework#144
