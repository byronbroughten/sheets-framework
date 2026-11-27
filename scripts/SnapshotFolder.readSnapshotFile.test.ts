import { writeFileSync } from "node:fs";
import { join } from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import { snapshotsFolder } from "./snapshotFixture.ts";
import { readSnapshotFile } from "./SnapshotFolder.ts";

afterEach(() => {
  vi.unstubAllEnvs();
});

const savedColumn = {
  tabTitle: "Value Types",
  header: "Text",
  columnId: "c:vty:text",
  values: ["Apple"],
};

function seededSnapshot(column: Record<string, unknown>): string {
  const name = "snapshot-seeded.json";
  const contents = {
    target: "dev",
    command: "chore overwrite",
    takenAt: "2026-11-25T12:00:00.000Z",
    plan: "",
    snapshot: { columns: [column], fingerprints: [] },
  };
  writeFileSync(join(snapshotsFolder(), name), JSON.stringify(contents));
  return name;
}

describe("readSnapshotFile", () => {
  it("refuses a file whose saved column has no tableId, naming version history as its undo", () => {
    const name = seededSnapshot(savedColumn);
    expect(() => readSnapshotFile(name)).toThrow(
      "snapshot-seeded.json predates the snapshot's Table IDs, so restore can't place its columns; the spreadsheet's version history undoes that send.",
    );
  });
  it("reads a file whose saved columns each carry a tableId", () => {
    const name = seededSnapshot({ ...savedColumn, tableId: "1234" });
    expect(readSnapshotFile(name).command).toBe("chore overwrite");
  });
});
