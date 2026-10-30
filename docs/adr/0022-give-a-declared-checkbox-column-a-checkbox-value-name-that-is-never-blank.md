# Give a declared checkbox column a `checkbox` value name that is never blank

An untouched checkbox holds nothing and an unticked one `FALSE`, yet the operator sees the same box, so a declared checkbox column gets the value name `checkbox`, typed exactly `boolean`, whose blank reads as unchecked. The fact is stated once, on the value name's schema, not repeated on every checkbox column, and the lowest tier that knows a value name applies it, so Raw keeps reporting the wire. The name is earned by declaration only: a column whose type menu says Checkbox becomes `checkbox`, while one on Automatic that holds `TRUE` stays `boolean`, keeps its blank and stays untyped, and only a `checkbox` column can be a selector or carry the checkbox column operator.

Source: byronbroughten/byro-repo#12
