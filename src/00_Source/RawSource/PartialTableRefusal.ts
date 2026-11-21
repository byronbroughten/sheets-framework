import type { LocalWriteOperation } from "./RawSource.ts";

// Google names only the request it refused, so the Source hands back the operation that sent it.
export class PartialTableRefusal extends Error {
  declare readonly operation: LocalWriteOperation;
  constructor(message: string, operation: LocalWriteOperation) {
    super(message);
    this.operation = operation;
  }
}
