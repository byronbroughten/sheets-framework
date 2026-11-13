import { describe, expect, it } from "vitest";

import type {
  SyncConfigs,
  SyncConflict,
  SyncDryRunResult,
  SyncRunResult,
  SyncSendResult,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { SyncRunReport } from "./SyncRunReport";

// The eight cases sheets-framework#216 retested on fresh agents; names and IDs are illustrative.

const noConfigChanges: SyncConfigs = {
  heldBack: [],
  renamed: [],
  namesMoved: [],
  rewrites: [],
  typedTraitChanges: [],
  notes: [],
};

const appDryRun = {
  mode: "dry run",
  target: "app",
  command: "npm run app:sync",
  generatedDir: "src/generated/",
  alterationsPath: "src/alterations.ts",
  checks: [],
  refusals: [],
  stops: [],
  plan: [],
  notices: [],
  configs: noConfigChanges,
} as const satisfies Omit<SyncDryRunResult, "outcome">;

const appSend = {
  ...appDryRun,
  mode: "send",
  landed: [],
  notSent: [],
  didntLand: [],
} as const satisfies Omit<SyncSendResult, "outcome">;

function render(result: SyncRunResult): string[] {
  return SyncRunReport.init(result).lines;
}

function resultLine(lines: string[]): string | undefined {
  return lines[1];
}

function sections(lines: string[]): string[] {
  return lines.flatMap((line) => {
    const match = /^== (.+) ==$/.exec(line);
    return match?.[1] === undefined ? [] : [match[1].split(":")[0] ?? ""];
  });
}

describe("SyncRunReport, case A: both changed", () => {
  const lines = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      {
        kind: "conflict",
        subject: { table: "occPayment", column: "amount", path: "columnType" },
        label: "both changed",
        cause:
          "A person changed this on the sheet after the last sync, and your entry changes it too.",
        yours: "CURRENCY",
        agreed: "AUTOMATIC",
        live: "NUMBER",
      },
    ],
    plan: [
      {
        subject: { table: "occPayment" },
        what: "set amount's type to CURRENCY (resends all 15 columns).",
        effects: [
          {
            kind: "dropdownReset",
            setting: "allow",
            columns: [
              "payerCategory",
              "nonOccupantPayerId",
              "formOfPayment",
              "paymentStatus",
            ],
          },
        ],
      },
    ],
    configs: {
      ...noConfigChanges,
      heldBack: [
        {
          subject: {
            table: "occPayment",
            column: "amount",
            path: "columnType",
          },
          value: "AUTOMATIC",
        },
      ],
    },
  });
  const text = lines.join("\n");

  it("leads with the target and mode, then says it stopped", () => {
    expect(lines[0]).toBe("sync · app · dry run");
    expect(resultLine(lines)).toBe(
      "STOPPED: 1 conflict. Nothing was sent, and nothing will be until it is resolved.",
    );
  });

  it("orders its sections", () => {
    expect(sections(lines)).toEqual(["Stops", "Plan", "Configs", "Next"]);
  });

  it("shows the three rows, the cause and every way out", () => {
    expect(text).toContain(
      "CONFLICT 1 · occPayment › amount › columnType · both changed",
    );
    expect(text).toMatch(/alterations\.ts \(yours\) +CURRENCY/);
    expect(text).toMatch(/configs \(last agreed\) +AUTOMATIC/);
    expect(text).toMatch(/live sheet +NUMBER/);
    expect(text).toContain("A person changed this on the sheet");
    expect(text).toContain(
      "adopt  Remove the entry, or set it to NUMBER. Free, but your change to this trait is given up.",
    );
    expect(text).toContain(
      "force  Send CURRENCY over the sheet's NUMBER. Needs the developer's yes for this trait; the yes covers one run.",
    );
    expect(text).toContain(
      "ask    Change nothing and ask the developer which value they want.",
    );
  });

  it("never prints force's flag", () => {
    expect(text).not.toContain("--force");
  });

  it("indents the dropdown reset under its send", () => {
    expect(text).toContain(
      ' 1. occPayment: set amount\'s type to CURRENCY (resends all 15 columns).\n    Resets the style and colours of 4 dropdowns: payerCategory, nonOccupantPayerId, formOfPayment, paymentStatus (dropdownStyleReset: "allow").',
    );
  });

  it("says what the configs keep while the conflict is open", () => {
    expect(text).toContain(
      "occPayment › amount › columnType: the configs keep AUTOMATIC for this trait until the conflict is resolved.",
    );
    expect(text).toContain("Resolve conflict 1, then rerun npm run app:sync.");
  });
});

describe("SyncRunReport, case B: a held overwrite of typed notes", () => {
  const lines = render({
    ...appDryRun,
    outcome: { kind: "ready" },
    held: {
      groups: [
        {
          subject: { table: "occPayment", column: "notes" },
          what: "the formula alteration overwrites entered data in 41 cells:",
          cells: 41,
          lines: [
            'opy-0007  "Paid by cheque #1182"  → formula',
            "… and 40 more in notes, counted, not listed",
          ],
        },
      ],
      token: "7f3a9c",
      fullListPath: ".probe/sync-held.json",
    },
    plan: [
      {
        subject: { table: "occPayment", column: "notes" },
        what: "write the formula into all 212 body rows. New rows copy it.",
        effects: [{ kind: "overwrites", cells: 41, blank: 171 }],
      },
    ],
  });
  const text = lines.join("\n");

  it("counts the held cells in the result line", () => {
    expect(resultLine(lines)).toBe(
      "READY, 41 HELD: --send sends only with --confirm (see Held). This was a dry run: nothing was sent.",
    );
  });

  it("orders its sections", () => {
    expect(sections(lines)).toEqual(["Held", "Plan", "Configs", "Next"]);
  });

  it("ends the held block with the full list, the token and the rerun", () => {
    expect(text).toContain(
      [
        "== Held: 41 cells, sent only with --confirm ==",
        "occPayment › notes  the formula alteration overwrites entered data in 41 cells:",
        '  opy-0007  "Paid by cheque #1182"  → formula',
        "  … and 40 more in notes, counted, not listed",
        "Full list, every cell with its value: .probe/sync-held.json",
        "Token 7f3a9c: binds this target and exactly this list.",
        "Rerun with --send --confirm 7f3a9c to send these.",
      ].join("\n"),
    );
    expect(text).toContain(
      "--send refuses this plan without --confirm 7f3a9c, because of the held items.",
    );
  });

  it("says the configs are unchanged", () => {
    expect(text).toContain("Refreshed from the live sheet. No changes.");
  });
});

describe("SyncRunReport, case C: a dropdown reset under ask, a rule added first, ID first", () => {
  const lines = render({
    ...appDryRun,
    target: "dev",
    command: "npm run dev:sync",
    generatedDir: "dev/generated/",
    alterationsPath: "dev/alterations.ts",
    outcome: { kind: "ready" },
    held: {
      groups: [
        {
          subject: { table: "item" },
          what: 'dropdown style reset (dropdownStyleReset: "ask"):',
          lines: [
            "Sending item's column list resets these dropdowns to plain chips and clears their option colours: status, category.",
          ],
        },
      ],
      token: "b21e07",
      fullListPath: "dev/.probe/sync-held.json",
    },
    plan: [
      {
        subject: { table: "item" },
        what: "set weight's type to NUMBER (resends all 9 columns).",
        effects: [
          {
            kind: "dropdownReset",
            setting: "ask",
            columns: ["status", "category"],
          },
        ],
      },
      {
        subject: { table: "item" },
        what: "add rule item·weight·NUMBER_LESS(0)·text#cc0000.",
        effects: [
          {
            kind: "addedFirst",
            outranks: ['item·all·=${status}="Retired"·bg#eeeeee'],
          },
        ],
      },
    ],
    notices: [{ kind: "idFirst", table: "item" }],
  });
  const text = lines.join("\n");

  it("leads with dev and counts the held item", () => {
    expect(lines[0]).toBe("sync · dev · dry run");
    expect(resultLine(lines)).toBe(
      "READY, 1 HELD: --send sends only with --confirm (see Held). This was a dry run: nothing was sent.",
    );
  });

  it("orders its sections", () => {
    expect(sections(lines)).toEqual([
      "Held",
      "Plan",
      "Notices",
      "Configs",
      "Next",
    ]);
  });

  it("indents each side effect under its send, in send order", () => {
    expect(text).toContain(
      [
        "== Plan: what --send sends, in order ==",
        " 1. item: set weight's type to NUMBER (resends all 9 columns).",
        '    Resets the style and colours of 2 dropdowns: status, category (dropdownStyleReset: "ask", so held, above).',
        " 2. item: add rule item·weight·NUMBER_LESS(0)·text#cc0000.",
        '    Added first; it outranks 1 rule on these cells: item·all·=${status}="Retired"·bg#eeeeee.',
      ].join("\n"),
    );
  });

  it("counts a held item that isn't cells as an item", () => {
    expect(text).toContain("== Held: 1 item, sent only with --confirm ==");
    expect(text).toContain("because of the held item.");
  });

  it("says why ID first matters", () => {
    expect(text).toContain(
      "item: ID is the first column. Sheets' Table name chip covers the cell above it, so item's Fill Row IDs checkbox can't be clicked.",
    );
  });
});

describe("SyncRunReport, case D: deleting a Table and a tab", () => {
  const lines = render({
    ...appDryRun,
    outcome: { kind: "ready" },
    checks: [
      {
        kind: "outsideReferences",
        subject: { table: "occChargeArchive" },
        effect: "blank",
        references: [
          { cell: "Ledger!C2", formula: "=SUM(occChargeArchive[Amount])" },
          { cell: "Ledger!C3", formula: "=COUNTA(occChargeArchive[ID])" },
          {
            cell: "Owner Report!F9",
            formula:
              "=SUMIFS(occChargeArchive[Amount], occChargeArchive[Unit ID], A9)",
          },
        ],
      },
      {
        kind: "outsideReferences",
        subject: { tab: "Old Imports" },
        effect: "ref",
        references: [{ cell: "Ledger!H1", formula: "='Old Imports'!B2" }],
      },
    ],
    held: {
      groups: [
        {
          subject: { table: "occChargeArchive" },
          what: "deleting the Table deletes entered data in 1,204 cells, over 318 rows:",
          cells: 1204,
          lines: ["description 318 · amount 318 · chargeDate 318 · unitId 250"],
        },
        {
          subject: { tab: "Old Imports" },
          what: "deleting the tab deletes entered data in importRaw (a Table without API access): 380 cells",
          cells: 380,
          lines: [],
        },
        {
          subject: { tab: "Old Imports" },
          what: "deleting the tab deletes 46 non-blank cells outside any Table, formulas and values",
          cells: 46,
          lines: [],
        },
      ],
      token: "4c9d10",
      fullListPath: ".probe/sync-held.json",
    },
    plan: [
      {
        subject: { table: "occChargeArchive" },
        what: "delete the Table, its cells and its head rows. Its tab, Charge Archive, stays, empty.",
        effects: [
          {
            kind: "rulesRemoved",
            handles: [
              "occChargeArchive·amount·NUMBER_LESS(0)·text#cc0000",
              "occChargeArchive·all·=${void}=TRUE·bg#eeeeee",
            ],
          },
          {
            kind: "rulePartlyOver",
            tab: "Charge Archive",
            range: "A2:H400",
            formula: "=$H2>1000",
          },
          { kind: "outsideReferences", count: 3, effect: "blank" },
        ],
      },
      {
        subject: { tab: "Old Imports", gid: 88123 },
        what: "delete the tab, with its Table importRaw.",
        effects: [{ kind: "outsideReferences", count: 1, effect: "ref" }],
      },
    ],
  });
  const text = lines.join("\n");

  it("counts both the held cells and the things to check", () => {
    expect(resultLine(lines)).toBe(
      "READY, 1,630 HELD: --send sends only with --confirm (see Held); 2 things to check (see Check before sending). This was a dry run: nothing was sent.",
    );
  });

  it("puts Check before sending first", () => {
    expect(sections(lines)).toEqual([
      "Check before sending",
      "Held",
      "Plan",
      "Configs",
      "Next",
    ]);
  });

  it("lists every outside formula the deletions break", () => {
    expect(text).toContain(
      "occChargeArchive: 3 formulas outside this Table read it. After the deletion they read blank cells, with no error:\n  Ledger!C2        =SUM(occChargeArchive[Amount])",
    );
    expect(text).toContain(
      "tab Old Imports: 1 formula on another tab reads it. After the deletion it shows #REF!:\n  Ledger!H1        ='Old Imports'!B2",
    );
  });

  it("shows the held cells' sum", () => {
    expect(text).toContain(
      "== Held: 1,630 cells: 1,204 + 380 + 46, sent only with --confirm ==",
    );
  });

  it("names the rules a deletion removes and points the outside formulas at the check", () => {
    expect(text).toContain(
      "    Removes 2 conditional-format rules wholly over it: occChargeArchive·amount·NUMBER_LESS(0)·text#cc0000, occChargeArchive·all·=${void}=TRUE·bg#eeeeee.",
    );
    expect(text).toContain(" 2. tab Old Imports (gid 88123): delete the tab");
    expect(text).toContain(
      "    1 formula elsewhere reads it and will show #REF! after the deletion (see Check before sending).",
    );
  });

  it("drops the unexplained prefix retirement", () => {
    expect(text).not.toContain("Retires the ID prefix");
  });
});

describe("SyncRunReport, case E: a rename and a swap the dry run took up", () => {
  const swapped = {
    table: "unit",
    column: "moveOut",
    rewrittenFrom: { wrote: "moveIn", columnId: "c:unt:k2PqA9x" },
  } as const;
  const lines = render({
    ...appDryRun,
    outcome: { kind: "ready" },
    checks: [
      {
        kind: "followedSwap",
        subject: { ...swapped, path: "emptyValueAllowed" },
        header: "Move out",
      },
    ],
    plan: [
      {
        subject: { table: "occCharge" },
        what: 'insert column "Late fee" (CURRENCY) before monthlyRent.',
        effects: [
          {
            kind: "rulesRewritten",
            table: "occCharge",
            count: 2,
            after: "insert",
          },
        ],
      },
      {
        subject: swapped,
        what: "emptyValueAllowed → false. Changes the configs only; nothing is sent to the sheet.",
        effects: [],
      },
    ],
    configs: {
      ...noConfigChanges,
      renamed: [
        {
          table: "occCharge",
          from: "rent",
          to: "monthlyRent",
          headerFrom: "Rent",
          headerTo: "Monthly rent",
        },
      ],
      namesMoved: [
        {
          table: "unit",
          column: "moveIn",
          wasId: "c:unt:k2PqA9x",
          nowId: "c:unt:Rb7sW0e",
        },
        {
          table: "unit",
          column: "moveOut",
          wasId: "c:unt:Rb7sW0e",
          nowId: "c:unt:k2PqA9x",
        },
      ],
      rewrites: [
        {
          from: 'occCharge.newColumns[0].before "rent"',
          to: '"monthlyRent"',
        },
        {
          from: "unit.columns.moveIn",
          to: "unit.columns.moveOut",
          why: 'c:unt:k2PqA9x, which a person retitled "Move out"',
        },
      ],
      notes: ["Not rewritten: endpoint code. tsc fails where they say rent."],
    },
  });
  const text = lines.join("\n");

  it("keeps nothing held from reading as a green light", () => {
    expect(resultLine(lines)).toBe(
      "READY: nothing held; 1 thing to check (see Check before sending). This was a dry run: nothing was sent.",
    );
    expect(text).toContain(
      "Nothing held, but read Check before sending first: npm run app:sync -- --send sends this plan.",
    );
  });

  it("orders its sections", () => {
    expect(sections(lines)).toEqual([
      "Check before sending",
      "Plan",
      "Configs",
      "Next",
    ]);
  });

  it("shows the rewritten key as written", () => {
    expect(text).toContain(
      'unit › moveOut (you wrote moveIn; c:unt:k2PqA9x) › emptyValueAllowed: sync rewrote this entry across a swap, so it still names the column it named before, now headed "Move out".',
    );
    expect(text).toContain(
      " 2. unit › moveOut (you wrote moveIn; c:unt:k2PqA9x): emptyValueAllowed → false.",
    );
  });

  it("indents the rule rewrites under the insert", () => {
    expect(text).toContain(
      "    Rewrites 2 conditional-format rules over occCharge after the insert, since Sheets leaves them wrong.",
    );
  });

  it("lists renames, names that moved and rewrites in the configs", () => {
    expect(text).toContain(
      '  occCharge › rent → monthlyRent    header "Rent" → "Monthly rent"',
    );
    expect(text).toContain(
      "Reread src/generated/idNameMap.ts before writing another entry:\n  unit › moveIn    was c:unt:k2PqA9x, now c:unt:Rb7sW0e",
    );
    expect(text).toContain(
      '  unit.columns.moveIn → unit.columns.moveOut    (c:unt:k2PqA9x, which a person retitled "Move out")',
    );
  });
});

describe("SyncRunReport's intended-state refusals", () => {
  const text = render({
    ...appDryRun,
    outcome: { kind: "refused" },
    refusals: [
      {
        kind: "touchesDeletion",
        deleted: { tab: "log" },
        deletedAt: "tabs.log",
        touchedBy: "tabs.item.before",
      },
    ],
  }).join("\n");

  it("names the deletion and what touches it", () => {
    expect(text).toContain("REFUSED 1 · tab log");
    expect(text).toContain(
      "tabs.log deletes tab log, and tabs.item.before touches it in the same run.",
    );
  });
});

describe("SyncRunReport's refusals before the read", () => {
  const text = render({
    ...appDryRun,
    outcome: { kind: "refused" },
    refusals: [
      { kind: "noSuchTable", entry: "tables", name: "vendor" },
      { kind: "noSuchTable", entry: "tabs", name: "vendors" },
      { kind: "notSendableYet", table: "item", path: "name.header" },
    ],
  }).join("\n");

  it("says a Table entry's Table is gone, under the entry", () => {
    expect(text).toContain("REFUSED 1 · tables › vendor");
    expect(text).toContain("  no such Table: adopt by removing the entry.");
  });

  it("says a tab entry's tab is gone, in words for a tab", () => {
    expect(text).toContain("REFUSED 2 · tabs › vendors");
    expect(text).toContain("  no such tab: adopt by removing the entry.");
  });

  it("names the Table and the trait it can't send yet", () => {
    expect(text).toContain("REFUSED 3 · item › name.header");
    expect(text).toContain(
      "  sync can't send name.header yet: for now it sends only a Table's letApiAccess and a column's emptyValueAllowed.",
    );
  });
});

describe("SyncRunReport's force way out", () => {
  function conflictText(
    subject: SyncConflict["subject"],
    gone: Pick<SyncConflict, "isColumnGone"> = {},
  ): string {
    return render({
      ...appDryRun,
      outcome: { kind: "stopped" },
      stops: [
        {
          kind: "conflict",
          subject,
          label: "both changed",
          cause: "Both changed it.",
          yours: "a",
          agreed: "b",
          live: "c",
          ...gone,
        },
      ],
    }).join("\n");
  }

  it("isn't offered for a column a person deleted, which has nothing to send to", () => {
    const subject = { table: "item", column: "name", path: "header" };
    expect(conflictText(subject, { isColumnGone: true })).not.toContain(
      "force",
    );
    expect(
      conflictText(
        { table: "item", column: "count", path: "formula" },
        { isColumnGone: true },
      ),
    ).not.toContain("force");
  });

  it("is offered for a column's trait", () => {
    expect(
      conflictText({ table: "item", column: "name", path: "header" }),
    ).toContain("force");
  });

  it("isn't offered for a tab's trait, which --force can't address", () => {
    const text = conflictText({ tab: "log", path: "title" });
    expect(text).toContain("tab log › title");
    expect(text).not.toContain("force");
  });
});

describe("SyncRunReport, case F: three refusals, then a head-row stop", () => {
  const ruleHandle = "occCharge·all·=${void}=TRUE·bg#eeeeee";
  const refused = render({
    ...appDryRun,
    outcome: { kind: "refused" },
    refusals: [
      { kind: "besideGrant", table: "vendor", path: "newColumns" },
      {
        kind: "noSuchRule",
        table: "occCharge",
        path: `conditionalFormats.replace["${ruleHandle}"]`,
        handles: [
          "occCharge·all·=${void}=TRUE·bg#f4cccc",
          "occCharge·amount·NUMBER_LESS(0)·text#cc0000",
        ],
      },
      {
        kind: "unknownBrace",
        table: "occCharge",
        path: "conditionalFormats.add[0].formula",
        brace: "${latefee}",
        columns: ["id", "amount", "lateFee", "void"],
      },
    ],
  });
  const refusedText = refused.join("\n");
  const stopped = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      { kind: "headRows", table: "vendor", tab: "Vendors", headerRow: 6 },
    ],
  });
  const stoppedText = stopped.join("\n");

  it("refuses before reading, with no configs refresh", () => {
    expect(resultLine(refused)).toBe(
      "REFUSED BEFORE READING: 3 entries can't be planned. Nothing was read or sent; alterations.ts is unchanged.",
    );
    expect(sections(refused)).toEqual(["Refused", "Next"]);
  });

  it("says the grant takes its own run, which other Tables' entries may share", () => {
    expect(refusedText).toContain("REFUSED 1 · vendor › newColumns");
    expect(refusedText).toContain(
      "  Granting access takes its own run: keep letApiAccess: true on vendor, take this entry out, and run sync.\n  Other Tables' entries may share that run.",
    );
  });

  it("says to grant first for an entry on a Table without API access", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [{ kind: "noApiAccess", table: "vendor", path: "name" }],
    }).join("\n");
    expect(text).toContain("REFUSED 1 · vendor › name");
    expect(text).toContain("grant `letApiAccess` first.");
  });

  it("lists the Table's current rule handles under no such rule", () => {
    expect(refusedText).toContain(
      "Adopt by removing the entry.\n  occCharge's rules now, by handle:\n    occCharge·all·=${void}=TRUE·bg#f4cccc\n    occCharge·amount·NUMBER_LESS(0)·text#cc0000",
    );
  });

  it("names a brace's Table columns", () => {
    expect(refusedText).toContain(
      "${latefee} names no column of occCharge. Its columns: id, amount, lateFee, void.",
    );
    expect(refusedText).toContain(
      "Fix each refused entry, then rerun npm run app:sync.",
    );
  });

  it("stops when a grant's Table can't take head rows on row 4", () => {
    expect(resultLine(stopped)).toBe(
      "STOPPED: 1 Table can't take head rows. Nothing was sent, and nothing will be until it is resolved.",
    );
    expect(sections(stopped)).toEqual(["Stops", "Configs", "Next"]);
    expect(stoppedText).toContain(
      "STOP 1 · vendor › letApiAccess\n  vendor's header row is on row 6 of its tab, Vendors.",
    );
    expect(stoppedText).toContain(
      "Resolve stop 1, then rerun npm run app:sync.",
    );
  });
});

describe("SyncRunReport, case G: a send stopped partway", () => {
  const lines = render({
    ...appSend,
    outcome: {
      kind: "stoppedPartway",
      landedFlushes: 1,
      flushCount: 2,
      heldBy: "mismatched",
    },
    landed: [
      {
        kind: "createdColumn",
        table: "occCharge",
        header: "Late fee",
        column: "lateFee",
        columnId: "c:och:Q8x1mZp",
        columnType: "CURRENCY",
        position: "at the Table's end",
      },
    ],
    notSent: [
      {
        subject: { table: "occCharge", column: "notes" },
        what: "the formula alteration.",
      },
    ],
    configs: {
      ...noConfigChanges,
      notes: ["Taken after flush 1: occCharge now has lateFee."],
    },
  });
  const text = lines.join("\n");

  it("says which flush landed and which was refused", () => {
    expect(lines[0]).toBe("sync · app · send");
    expect(resultLine(lines)).toBe(
      "STOPPED PARTWAY: flush 1 of 2 landed; flush 2 was refused, because the sheet or the plan changed after the dry run, so --confirm doesn't cover its held items. Nothing after it was sent.",
    );
  });

  it("orders its sections", () => {
    expect(sections(lines)).toEqual(["Landed", "Not sent", "Configs", "Next"]);
  });

  it("prints a created column's codebase name and ID", () => {
    expect(text).toContain(
      'occCharge: created column "Late fee" as lateFee, c:och:Q8x1mZp, CURRENCY, at the Table\'s end.',
    );
    expect(text).toContain("occCharge › notes: the formula alteration.");
    expect(text).not.toContain("alterations.ts is now empty.");
  });

  it("prints no held list and no token, only the rerun", () => {
    expect(text).not.toContain("Token");
    expect(text).toContain("Rerun the dry run: npm run app:sync");
  });
});

describe("SyncRunReport, a send refused at its first flush", () => {
  const lines = render({
    ...appSend,
    outcome: { kind: "stoppedBeforeSending", heldBy: "uncovered" },
    landed: [],
    notSent: [
      {
        subject: { table: "occCharge", column: "notes" },
        what: "the formula alteration.",
      },
    ],
    configs: noConfigChanges,
  });
  const text = lines.join("\n");

  it("says nothing was sent, never that it stopped partway", () => {
    expect(resultLine(lines)).toBe(
      "STOPPED: the send was refused before any flush landed, because it holds items and the send had no --confirm. Nothing was sent.",
    );
    expect(text).not.toContain("PARTWAY");
  });

  it("points to the dry run", () => {
    expect(text).toContain("Rerun the dry run: npm run app:sync");
  });
});

describe("SyncRunReport, case H: a clean send", () => {
  const lines = render({
    ...appSend,
    outcome: {
      kind: "sent",
      summary: "2 alterations landed; 3 blank IDs filled",
    },
    checks: [
      {
        kind: "handDrawnRule",
        table: "occPayment",
        tab: "Occ Payment",
        range: "A3:N3 (it covers a head row)",
        formula: '=$H3=""',
        handle: 'occPayment·all·=$H3=""·bg#d9ead3',
        cause:
          "couldn't be rewritten, and the move left it pointing at the wrong column",
      },
    ],
    landed: [
      {
        kind: "send",
        subject: { table: "occPayment", column: "processed" },
        what: "moved before notes.",
      },
      {
        kind: "createdColumn",
        table: "occPayment",
        header: "Receipt no.",
        column: "receiptNo",
        columnId: "c:opy:Xw3nB7d",
        columnType: "TEXT",
        position: "at the Table's end",
      },
    ],
    plan: [
      {
        subject: { table: "occPayment" },
        what: "rewrote 3 conditional-format rules after the move, since Sheets leaves them wrong.",
        effects: [
          {
            kind: "ruleRespanned",
            handle: "occPayment·all·=${processed}=FALSE·bg#fff2cc",
            fromRow: 127,
            toRow: 212,
          },
        ],
      },
      {
        subject: { table: "household" },
        what: "filled 3 blank IDs: hh-8Kd2Lq0, hh-T41xZpa, hh-c9WmE2r (rows 140 to 142).",
        effects: [],
      },
    ],
    configs: {
      ...noConfigChanges,
      typedTraitChanges: [
        {
          table: "household",
          column: "phone",
          trait: "columnType",
          from: "TEXT",
          to: "NUMBER",
        },
      ],
    },
  });
  const text = lines.join("\n");

  it("says it sent, counting the thing to check", () => {
    expect(resultLine(lines)).toBe(
      "SENT: 2 alterations landed; 3 blank IDs filled; 1 thing to check (see Check before sending).",
    );
  });

  it("orders its sections", () => {
    expect(sections(lines)).toEqual([
      "Check before sending",
      "Landed",
      "Plan",
      "Configs",
      "Next",
    ]);
  });

  it("prints the hand-drawn rule's handle so a remove can name it", () => {
    expect(text).toContain('Its handle: occPayment·all·=$H3=""·bg#d9ead3.');
  });

  it("says alterations.ts is empty once everything landed", () => {
    expect(text).toContain(
      'occPayment: created column "Receipt no." as receiptNo, c:opy:Xw3nB7d, TEXT, at the Table\'s end.\nalterations.ts is now empty.',
    );
  });

  it("prints the respanned rule's handle and the typed trait a person changed", () => {
    expect(text).toContain(
      "Re-spans 1 rule that ended at row 127 to the Table's last body row, 212: occPayment·all·=${processed}=FALSE·bg#fff2cc.",
    );
    expect(text).toContain("  household › phone    columnType TEXT → NUMBER");
    expect(text).toContain(
      "sync wrote src/generated/ and src/alterations.ts. Run npm run check, then commit them with your change.",
    );
  });
});
