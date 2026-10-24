# Carry a selector's options inside one selector object

An endpoint's selector is an object holding its checkbox column and its options, `retainSelection` and `requireOneRow`, rather than a bare column name with flags beside it on the entry, so an endpoint with no selector has nowhere to write an option and the compiler rejects what would otherwise be a silent no-op. The bare-name shorthand was not kept as a union alternative, since two shapes would make every read site normalize first; each option is a default-false directive, so the common case, a selection that clears after a successful run, needs no declaration.

Source: byronbroughten/byro-repo#11, byronbroughten/byro-repo#18
