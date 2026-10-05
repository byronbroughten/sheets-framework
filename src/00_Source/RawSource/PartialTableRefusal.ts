import type { LocalWriteOperation } from "./RawSource";

// Google names only the request it refused, so the Source hands back the operation that sent it.
export class PartialTableRefusal extends Error {
  constructor(
    message: string,
    readonly operation: LocalWriteOperation,
  ) {
    super(message);
  }
}
