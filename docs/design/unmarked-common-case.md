# Give the common case the unmarked name

Design reasoning. The one-line principle lives in [`docs/design.md`](../design.md); this file holds the argument and its instances.

## The argument

The name a reader reaches for by reflex should be the one they want most of the time. Make the common case free and let the rare case cost exactly one word — and decide which case is common by counting call sites rather than by guessing.

## Instances

A census showed roughly twice as many call sites reaching for a sheet's or column's data view as for its metadata view, yet the metadata view held the unmarked name and every data chain paid a `.data` hop; swapping primacy made the common case free and the rare case one word (#6). Nearly every cell read is one where a blank means the endpoint cannot do its job, yet `value` handed back `""` typed into the union and the safe read was the longer `valueNotEmpty` — inverting that made the reflexive read the correct one and left `valueOrEmpty` for the caller who has decided what blank means (#8).

## Corollaries

Where "the common case" differs per subject, let the subject declare it rather than picking one meaning for all of them. `value` throwing on a blank was right for most columns and wrong for the ones that are legitimately blank, and no call site could tell which kind it was looking at — so the column's own **Empty value allowed** box now decides what its unmarked read means, and `valueNotEmpty` came back as the marked word for a call site stricter than its column (#13). That carries the principle further rather than reversing #8: the unmarked word is the reflexive correct read on every column now, instead of only on one that cannot be blank.

The marked name is a record of intent, not merely a longer spelling. `valueOrEmpty` at a call site says blank was thought about, which is a fact a later reader cannot recover from the surrounding logic.

A tier that shouldn't offer the common case declines the unmarked word rather than reusing it for something else. Raw exposes only `valueOrEmpty`, and Identified gave `value` up once the word came to mean "whatever this column declared" — a declaration Identified cannot see, since it resolves columns by generated id (#13). So `value` means one thing everywhere it exists and moving a call between tiers can't silently change its failure mode.
