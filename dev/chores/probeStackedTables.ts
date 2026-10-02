import { writeFileSync } from "node:fs";

import { SheetsTransport } from "../../scripts/nodeHost";
import type { Chore } from "../../src/chores/Chore";
import { devSpreadsheetId } from "./devFixtures/devFixtureSheets";

// Throwaway probe for sheets-framework#56; delete after the run is recorded.
const probe = {
  outPath:
    "/private/tmp/claude-501/-Users-byronbroughten-Documents-Code-byro-repo/3004effa-c97c-439b-b1fd-a2028c63a26f/scratchpad/probe56.txt",
  base: "https://sheets.googleapis.com/v4/spreadsheets",
  firstGid: 9900001,
} as const;

type Req = Record<string, unknown>;
interface Step {
  label: string;
  requests: Req[];
}
interface Case {
  key: string;
  setup: (gid: number) => Req[];
  steps: (gid: number) => Step[];
}

function cells(gid: number, r: number, c: number, rows: string[][]): Req {
  return {
    updateCells: {
      start: { sheetId: gid, rowIndex: r, columnIndex: c },
      rows: rows.map((row) => ({
        values: row.map((v) => (v ? { userEnteredValue: { stringValue: v } } : {})),
      })),
      fields: "userEnteredValue",
    },
  };
}
function gridRange(gid: number, r0: number, r1: number, c0: number, c1: number): Req {
  return { sheetId: gid, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}
// Header, body and the three bookkeeping rows above, each cell labelled with its Table's letter.
function tableWithRows(gid: number, id: string, t: string, header: number, bodyRows: number, c0: number, width: number, bookkeeping = true): Req[] {
  const row = (tag: string) => Array.from({ length: width }, (_, i) => `${t}${tag}${i}`);
  const reqs: Req[] = [];
  if (bookkeeping) reqs.push(cells(gid, header - 3, c0, [row("id"), row("gp"), row("ac")]));
  reqs.push(cells(gid, header, c0, [row("h")]));
  reqs.push(cells(gid, header + 1, c0, Array.from({ length: bodyRows }, (_, b) => row(`${b}_`))));
  reqs.push({ addTable: { table: { tableId: id, name: id, range: gridRange(gid, header, header + 1 + bodyRows, c0, c0 + width) } } });
  return reqs;
}
function append(gid: number, id: string, n: number): Req {
  return {
    appendCells: {
      sheetId: gid,
      tableId: id,
      rows: Array.from({ length: n }, (_, i) => ({
        values: Array.from({ length: 3 }, () => ({ userEnteredValue: { stringValue: `new${i}` } })),
      })),
      fields: "userEnteredValue",
    },
  };
}
function insertCols(gid: number, r0: number, r1: number, c: number): Req {
  return { insertRange: { range: gridRange(gid, r0, r1, c, c + 1), shiftDimension: "COLUMNS" } };
}
function deleteRows(gid: number, r0: number, r1: number, c0: number, c1: number): Req {
  return { deleteRange: { range: gridRange(gid, r0, r1, c0, c1), shiftDimension: "ROWS" } };
}

// Item 1: A at A..C, header row 3, three body rows (4..6), so its next row down is 7.
function stackCase(key: string, below: (gid: number) => Req[], n: number): Case {
  const a = `P56_${key}_A`;
  return {
    key,
    setup: (gid) => [...tableWithRows(gid, a, "A", 3, 3, 0, 3), ...below(gid)],
    steps: (gid) => [{ label: `append ${n} row(s) to A`, requests: [append(gid, a, n)] }],
  };
}

// Item 2: A at B..D, header row 5, body 6..8; the insert band is rows 2..8 (header −3 to A's last row).
type Variant = "inside" | "past" | "pastWiden";
function columnCase(key: string, variant: Variant, b?: { header: number; body: number; bookkeeping: boolean }): Case {
  const a = `P56_${key}_A`;
  return {
    key,
    setup: (gid) => [
      ...tableWithRows(gid, a, "A", 5, 3, 1, 3),
      cells(gid, 0, 0, [["L0", "", "", "UP", "UP", "UP", "UP"], ["", "", "", "", "UP", "", "UP"]]),
      cells(gid, 9, 0, [["L9", "", "", "DN", "DN", "DN", "DN", "DN"]]),
      cells(gid, 10, 4, [["DN", "", "DN"]]),
      ...(b
        ? tableWithRows(gid, `P56_${key}_B`, "B", b.header, b.body, 6, 3, b.bookkeeping)
        : [cells(gid, 2, 4, [["R2"]]), cells(gid, 5, 5, [["R5"]]), cells(gid, 8, 6, [["R8"]])]),
    ],
    steps: (gid) => {
      if (variant === "inside") return [{ label: "insertRange at A's last column (D), rows 2..8", requests: [insertCols(gid, 2, 9, 3)] }];
      if (variant === "past") return [{ label: "insertRange just past A (E), rows 2..8", requests: [insertCols(gid, 2, 9, 4)] }];
      return [
        {
          label: "insertRange just past A (E), rows 2..8, + updateTable A to B..E",
          requests: [insertCols(gid, 2, 9, 4), { updateTable: { table: { tableId: a, range: gridRange(gid, 5, 9, 1, 5) }, fields: "range" } }],
        },
      ];
    },
  };
}

const cases: Case[] = [
  stackCase("1base", () => [], 1),
  stackCase("1a1", (g) => [cells(g, 7, 0, [["X"]])], 1),
  stackCase("1a3", (g) => [cells(g, 8, 0, [["X"]])], 3),
  stackCase("1aOut", (g) => [cells(g, 7, 4, [["X"]])], 1),
  stackCase("1b1", (g) => tableWithRows(g, "P56_1b1_B", "B", 10, 2, 0, 3), 1),
  stackCase("1b3", (g) => tableWithRows(g, "P56_1b3_B", "B", 11, 2, 0, 3), 3),
  stackCase("1c1", (g) => tableWithRows(g, "P56_1c1_B", "B", 7, 2, 0, 3, false), 1),
  stackCase("1c3", (g) => tableWithRows(g, "P56_1c3_B", "B", 8, 2, 0, 3, false), 3),
  columnCase("2in", "inside"),
  columnCase("2past", "past"),
  columnCase("2pw", "pastWiden"),
  columnCase("2eqIn", "inside", { header: 5, body: 3, bookkeeping: true }),
  columnCase("2eqPw", "pastWiden", { header: 5, body: 3, bookkeeping: true }),
  columnCase("2shortIn", "inside", { header: 5, body: 2, bookkeeping: true }),
  columnCase("2tallIn", "inside", { header: 5, body: 6, bookkeeping: false }),
  columnCase("2tallBkIn", "inside", { header: 5, body: 6, bookkeeping: true }),
  columnCase("2tallPw", "pastWiden", { header: 5, body: 6, bookkeeping: true }),
  columnCase("2highIn", "inside", { header: 1, body: 7, bookkeeping: false }),
  {
    key: "3one",
    setup: (g) => [...tableWithRows(g, "P56_3one_A", "A", 3, 1, 0, 3), cells(g, 5, 0, [["BELOW"]]), cells(g, 4, 4, [["SIDE"]])],
    steps: (g) => [
      { label: "deleteRange A's only body row (row 4, A..C)", requests: [deleteRows(g, 4, 5, 0, 3)] },
      { label: "append 1 row to A", requests: [append(g, "P56_3one_A", 1)] },
    ],
  },
  {
    key: "3two",
    setup: (g) => [...tableWithRows(g, "P56_3two_A", "A", 3, 2, 0, 3), cells(g, 6, 0, [["BELOW"]])],
    steps: (g) => [{ label: "deleteRange both body rows (rows 4..5, A..C)", requests: [deleteRows(g, 4, 6, 0, 3)] }],
  },
];

export const probeStackedTables: Chore = {
  description: "Throwaway probe for sheets-framework#56 (dev only, --send only): stacked growth, Table-bounded column inserts, last-row delete.",
  action: (_ss, { spreadsheetId }) => {
    if (spreadsheetId !== devSpreadsheetId) throw new Error("dev spreadsheet only");
    if (!process.argv.includes("--send")) return "Dry run: this probe writes directly, so it does nothing without --send.";
    const transport = SheetsTransport.init();
    const t = {
      send: (request: Parameters<SheetsTransport["send"]>[0]): unknown => {
        for (let attempt = 0; ; attempt++) {
          try {
            return transport.send(request);
          } catch (e) {
            if (attempt > 6 || !String((e as Error).message).includes(" 429 ")) throw e;
            Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 20000);
          }
        }
      },
    };
    const get = (query: string) => t.send({ method: "GET", url: `${probe.base}/${spreadsheetId}?${query}`, body: null }) as any;
    const batch = (requests: Req[]): string => {
      try {
        t.send({ method: "POST", url: `${probe.base}/${spreadsheetId}:batchUpdate`, body: JSON.stringify({ requests }) });
        return "OK";
      } catch (e) {
        const msg = String((e as Error).message);
        const m = msg.match(/"message":\s*"((?:[^"\\]|\\.)*)"/);
        return `ERROR ${m ? m[1] : msg.slice(0, 400)}`;
      }
    };
    const sheetGids = () => (get("fields=sheets(properties(sheetId))").sheets as any[]).map((s) => s.properties.sheetId as number);
    const gids = cases.map((_, i) => probe.firstGid + i);
    const removeProbeSheets = () => {
      const live = sheetGids();
      const present = gids.filter((g) => live.includes(g));
      return present.length ? batch(present.map((sheetId) => ({ deleteSheet: { sheetId } }))) : "nothing to remove";
    };
    removeProbeSheets();

    const out: string[] = [];
    const a1 = (r: any) => `${col(r.startColumnIndex ?? 0)}${r.startRowIndex ?? 0}:${col((r.endColumnIndex ?? 1) - 1)}${(r.endRowIndex ?? 1) - 1}`;
    const snapshot = (title: string) => {
      const fields = encodeURIComponent("sheets(tables(name,range),data(rowData(values(formattedValue))))");
      const s = get(`ranges=${encodeURIComponent(`'${title}'!A1:K14`)}&includeGridData=true&fields=${fields}`).sheets[0];
      const tables = (s.tables ?? []).map((tb: any) => `${String(tb.name).replace(/^P56_[^_]+_/, "")}=${a1(tb.range)}`).join("  ");
      out.push(`    tables (0-based rows): ${tables || "none"}`);
      out.push(`         ${Array.from({ length: 11 }, (_, i) => col(i).padEnd(6)).join("")}`);
      (s.data?.[0]?.rowData ?? []).forEach((rd: any, r: number) => {
        const vals = Array.from({ length: 11 }, (_, c) => String(rd.values?.[c]?.formattedValue ?? ".").padEnd(6));
        out.push(`    ${String(r).padStart(3)}  ${vals.join("")}`);
      });
    };
    cases.forEach((c, i) => {
      const gid = gids[i];
      const title = `P56 ${c.key}`;
      out.push(`== ${c.key}`);
      const setup = batch([{ addSheet: { properties: { sheetId: gid, title, gridProperties: { rowCount: 30, columnCount: 14 } } } }, ...c.setup(gid)]);
      out.push(`  setup: ${setup}`);
      if (setup !== "OK") return;
      out.push("  before:");
      snapshot(title);
      c.steps(gid).forEach((step) => {
        out.push(`  step: ${step.label} -> ${batch(step.requests)}`);
        snapshot(title);
      });
    });
    const cleanup = removeProbeSheets();
    const left = (get("fields=sheets(properties(title),tables(name))").sheets as any[]).map((s) => `${s.properties.title}[${(s.tables ?? []).map((tb: any) => tb.name).join(",")}]`);
    out.push(`== cleanup: ${cleanup}; remaining: ${left.join(" ")}`);
    writeFileSync(probe.outPath, out.join("\n"));
    return `Wrote ${probe.outPath}`;
  },
};

function col(i: number): string {
  return String.fromCharCode(65 + i);
}
