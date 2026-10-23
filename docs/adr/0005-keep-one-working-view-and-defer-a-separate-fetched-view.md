# Keep one working view, and defer a separate fetched view

The state revamp split each tier's state into a working view, a fetch queue and a write queue, where the working view is the fetched sheet plus every queued write, with row indexes at their pre-flush positions. Splitting it again, into a fetched view that stays put until the flush and a working view that moves at once, was left out of scope, because the revamp changed how state is named and grouped, not what the working view does.

Source: byronbroughten/byro-repo#55, byronbroughten/byro-repo#58
