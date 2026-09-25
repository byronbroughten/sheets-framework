# Rules for `src/06_API/`

- **`EndpointRun` is built from `Api`'s `SpreadsheetNamedProps`, never a no-arg `init()`**, which mints fresh state and re-fetches everything `Api` already read.
- **Keep the endpoint selector an anonymous nested object and the run's prop type `EndpointDispatched`.** A named type breaks `Api.ts`'s assignment, not the file you edited.
- **In a run, a success-only step goes inside the `try` after the action; a precondition that must fail loudly goes inside the `try` before it.** No success flag.
- **A failure path discards queued changes before it writes status**, or the `finally` flush ships the half-finished run.
- Mechanics: [`docs/architecture/endpoint-dispatch.md`](../../docs/architecture/endpoint-dispatch.md).
