# Carry a run's outcome in a cell colour nothing reads back, and let a conditional format mask it

A run's outcome is a cell's background colour, light yellow while running, then green or red, and the read path never fetches it, because the audience is a person looking at the sheet and no code needs the outcome back. A conditional format outranks that painted colour, so a rule over a feedback column can hide the yellow "never reported back" state; that overlap is accepted rather than refused, because the developer controls which rules they add. Run states were not turned into rules either, because a run state is a fact about one run, not a condition on a cell's value.

Source: byronbroughten/byro-repo#4, byronbroughten/byro-repo#45, commit byronbroughten/byro-repo@ac7a7959
