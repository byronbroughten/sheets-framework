# Type a fixed-list dropdown column as the literal union its list declares

A dropdown column whose value source is a fixed list in Value Config is typed as the union of that list's members rather than as `string`, so code that routes on it, such as the build-ledger endpoint's switch over a charge reduction's two descriptions, is exhaustive and a member added to the sheet is a compile error at the switch rather than a line that renders wrong. That holds only because `makeValueConfigs` takes a `const` type parameter, at about 41k type instantiations; without it every member list widens to `string[]` with no error anywhere.

Source: byronbroughten/byro-repo#18
