import { describe, expect, it } from "vitest";

import type {
  SyncConfigs,
  SyncConflict,
  SyncDryRunResult,
  SyncRunResult,
  SyncSendResult,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import { SyncRunReport } from "./SyncRunReport.ts";

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
  alreadyMet: [],
  forcedTraits: [],
  notices: [],
  configs: noConfigChanges,
} as const satisfies Omit<SyncDryRunResult, "outcome">;

const appSend = {
  ...appDryRun,
  mode: "send",
  landed: [],
  notSent: [],
  didntLand: [],
  hasEmptiedAlterations: false,
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
      "force  Send CURRENCY over the sheet's NUMBER. Needs the developer's yes for this trait; the yes covers one run. Flag: --force occPayment/amount/columnType",
    );
    expect(text).toContain(
      "ask    Change nothing and ask the developer which value they want.",
    );
  });

  it("indents the dropdown reset under its send", () => {
    expect(text).toContain(
      " 1. occPayment: set amount's type to CURRENCY (resends all 15 columns).\n    Resets the style and colours of 4 dropdowns: payerCategory, nonOccupantPayerId, formOfPayment, paymentStatus (dropdownStyleReset: \"allow\"). This is a quirk of Google's API: setting any column's type resends the Table's column list, which resets every dropdown on it.",
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
      forcedGroups: [],
      block: {
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
      forcedGroups: [],
      block: {
        groups: [
          {
            subject: { table: "item" },
            what: 'dropdown style reset (dropdownStyleReset: "ask"):',
            isUndoneOnlyByVersionHistory: true,
            lines: [
              "Sending item's column list resets these dropdowns to plain chips and clears their option colours: status, category.",
            ],
          },
        ],
        token: "b21e07",
        fullListPath: "dev/.probe/sync-held.json",
      },
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
        "    Resets the style and colours of 2 dropdowns: status, category (dropdownStyleReset: \"ask\", so held, above). This is a quirk of Google's API: setting any column's type resends the Table's column list, which resets every dropdown on it.",
        " 2. item: add rule item·weight·NUMBER_LESS(0)·text#cc0000.",
        '    Added first; it outranks 1 rule on these cells: item·all·=${status}="Retired"·bg#eeeeee.',
      ].join("\n"),
    );
  });

  it("names the dropdown reset as undone only by version history", () => {
    expect(text).toContain(
      [
        'item  dropdown style reset (dropdownStyleReset: "ask"):',
        "  undone only by the spreadsheet's version history: the send snapshot can't hold it",
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
      forcedGroups: [],
      block: {
        groups: [
          {
            subject: { table: "occChargeArchive" },
            what: "deleting the Table deletes entered data in 1,204 cells, over 318 rows:",
            cells: 1204,
            lines: [
              "description 318 · amount 318 · chargeDate 318 · unitId 250",
            ],
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
      {
        subject: { table: "occCharge" },
        what: 'append column "Note" (TEXT).',
        effects: [
          {
            kind: "rulesRewritten",
            table: "occCharge",
            count: 1,
            after: "append",
          },
        ],
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
        { kind: "table", name: "lease", wasId: "lse", nowId: "occ" },
        { kind: "tab", name: "leases", wasId: "101", nowId: "202" },
        {
          kind: "column",
          table: "unit",
          column: "moveIn",
          wasId: "c:unt:k2PqA9x",
          nowId: "c:unt:Rb7sW0e",
        },
        {
          kind: "column",
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

  it("says an append's rules are widened over the new columns", () => {
    expect(text).toContain(
      "    Widens 1 conditional-format rule over occCharge after the append, since Sheets leaves them short of the new columns.",
    );
  });

  it("lists renames, names that moved and rewrites in the configs", () => {
    expect(text).toContain(
      '  occCharge › rent → monthlyRent    header "Rent" → "Monthly rent"',
    );
    expect(text).toContain(
      [
        "Names that now point at a different ID than at the last sync. Reread src/generated/idNameMap.ts before writing another entry:",
        "  Table lease    was lse, now occ",
        "  tab leases    was 101, now 202",
        "  column unit › moveIn    was c:unt:k2PqA9x, now c:unt:Rb7sW0e",
      ].join("\n"),
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
      {
        kind: "touchesDeletion",
        deleted: { tab: "log" },
        deletedAt: "tabs.log",
        touchedBy: "tables.logEntry",
        takenTable: "logEntry",
      },
    ],
  }).join("\n");

  it("names the deletion and what touches it", () => {
    expect(text).toContain("REFUSED 1 · tab log");
    expect(text).toContain(
      "tabs.log deletes tab log, and tabs.item.before touches it in the same run.",
    );
  });

  it("says a tab's deletion already takes a Table another entry deletes", () => {
    expect(text).toContain(
      "tabs.log deletes tab log, and tables.logEntry touches it in the same run.\n  The tab's deletion already takes the Table logEntry.",
    );
    expect(text.match(/already takes/g)).toHaveLength(1);
  });
});

describe("SyncRunReport's refusals before the read", () => {
  const text = render({
    ...appDryRun,
    outcome: { kind: "refused" },
    refusals: [
      { kind: "noSuchTable", entry: "tables", name: "vendor" },
      { kind: "noSuchTable", entry: "tabs", name: "vendors" },
      {
        kind: "notSendableYet",
        entry: "tables",
        name: "item",
        path: "newColumns[0].columnType",
      },
      {
        kind: "notSendableYet",
        entry: "tabs",
        name: "item",
        path: "newTables[0].conditionalFormats.add[0].format.bold",
      },
      {
        kind: "notSendableYet",
        entry: "newTabs",
        name: "Empty",
        path: "newTabs[0].tables",
      },
    ],
  }).join("\n");

  it("says a Table entry's Table is gone, naming it as notSendableYet does", () => {
    expect(text).toContain("REFUSED 1 · Table vendor\n");
    expect(text).toContain("  no such Table: adopt by removing the entry.");
  });

  it("says a tab entry's tab is gone, in words for a tab", () => {
    expect(text).toContain("REFUSED 2 · tab vendors\n");
    expect(text).toContain("  no such tab: adopt by removing the entry.");
  });

  it("names the Table and the trait it can't send yet", () => {
    expect(text).toContain("REFUSED 3 · Table item › newColumns[0].columnType");
    expect(text).toContain(
      "  sync can't send newColumns[0].columnType yet: for now it sends only a Table's name, deleted, letApiAccess, newColumns, moves, conditionalFormats.remove, conditionalFormats.replace, conditionalFormats.add, a column's header, columnType, formula, deleted, emptyValueAllowed and a tab's title, color, before, frozenRows, frozenColumns, hidden, deleted, newTables.",
    );
  });

  it("names a tab's trait apart from a Table's of the same name", () => {
    expect(text).toContain(
      "REFUSED 4 · tab item › newTables[0].conditionalFormats.add[0].format.bold",
    );
  });

  it("names a new tab by its title", () => {
    expect(text).toContain("REFUSED 5 · new tab Empty › newTabs[0].tables");
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

  it("leaves adopt only removing the entry for a deleted column, which has no live value to take", () => {
    const text = conflictText(
      { table: "item", path: "moves[0]" },
      { isColumnGone: true },
    );
    expect(text).toContain(
      "adopt  Remove the entry. Free, but your change is given up.",
    );
    expect(text).not.toContain("set it to");
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

describe("SyncRunReport, case F: two refusals, then a head-row stop", () => {
  const refused = render({
    ...appDryRun,
    outcome: { kind: "refused" },
    refusals: [
      { kind: "besideGrant", table: "vendor", path: "newColumns" },
      {
        kind: "unknownBrace",
        table: "occCharge",
        path: "conditionalFormats.add[0].formula",
        brace: "${latefee}",
        braceTable: "occCharge",
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
      "REFUSED BEFORE READING: 2 entries can't be planned. Nothing was read or sent; alterations.ts is unchanged.",
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

  it("names a brace's Table columns", () => {
    expect(refusedText).toContain(
      "${latefee} names no column of occCharge. Its columns: id, amount, lateFee, void.",
    );
    expect(refusedText).toContain(
      "Fix each refused entry, then rerun npm run app:sync.",
    );
  });

  it("names a brace's missing Table, with no columns to list", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "unknownBraceTable",
          table: "occCharge",
          path: "amount.formula",
          brace: "${vendor[@cost]}",
          braceTable: "vendor",
        },
      ],
    }).join("\n");
    expect(text).toContain("REFUSED 1 · occCharge › amount.formula");
    expect(text).toContain(
      "${vendor[@cost]} names no Table: idNameMap.ts holds no Table named vendor.",
    );
  });

  it("names a rule's unknown column, with the columns it could name", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "unknownColumn",
          table: "occCharge",
          path: "conditionalFormats.add[0].columns",
          column: "lateFees",
          columns: ["id", "amount", "lateFee"],
        },
      ],
    }).join("\n");
    expect(text).toContain(
      "REFUSED 1 · occCharge › conditionalFormats.add[0].columns",
    );
    expect(text).toContain(
      "lateFees names no column of occCharge, nor one its newColumns create. Its columns: id, amount, lateFee.",
    );
  });

  it("names both newColumns headers that give one column name", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "repeatedNewColumn",
          table: "occCharge",
          path: "newColumns[1].header",
          firstPath: "newColumns[0].header",
          column: "lateFee",
        },
      ],
    }).join("\n");
    expect(text).toContain(
      "REFUSED 1 · occCharge › newColumns[1].header\n  newColumns[1].header gives the column name lateFee, as newColumns[0].header does: the two columns can't both be named it.\n  Give each a header whose column name is its own, then rerun.",
    );
  });

  it("names both the removal and the replace of one handle", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "removedAndReplaced",
          table: "occCharge",
          path: 'conditionalFormats.replace["amount | =TRUE"]',
          removedAt: 'conditionalFormats.remove["amount | =TRUE"]',
        },
      ],
    }).join("\n");
    expect(text).toContain(
      'REFUSED 1 · occCharge › conditionalFormats.replace["amount | =TRUE"]\n  conditionalFormats.remove["amount | =TRUE"] removes the rule this handle names, and conditionalFormats.replace["amount | =TRUE"] replaces it in the same run.\n  Keep one of the two, then rerun.',
    );
  });

  it("names both paths of each moves clash", () => {
    const clashes = [
      {
        clash: "sameBefore",
        path: "moves[1].before",
        otherPath: "moves[0].before",
        line: "moves[1].before puts its group before amount, as moves[0].before does: only one group can sit right before it.",
      },
      {
        clash: "repeatedColumn",
        path: "moves[1].columns[0]",
        otherPath: "moves[0].columns[0]",
        line: "moves[1].columns[0] moves amount, as moves[0].columns[0] does: a column can sit in only one group.",
      },
      {
        clash: "movedBefore",
        path: "moves[1].before",
        otherPath: "moves[0].columns[0]",
        line: "moves[1].before puts its group before amount, which moves[0].columns[0] moves: a group's target can't itself move.",
      },
    ] as const;
    for (const { clash, path, otherPath, line } of clashes) {
      const text = render({
        ...appDryRun,
        outcome: { kind: "refused" },
        refusals: [
          {
            kind: "clashingPlacements",
            clash,
            table: "occCharge",
            path,
            otherPath,
            column: "amount",
          },
        ],
      }).join("\n");
      expect(text).toContain(
        `REFUSED 1 · occCharge › ${path}\n  ${line}\n  Keep one of the two, then rerun.`,
      );
    }
  });

  it("says to name a renamed column by its old name this run", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "renamedColumn",
          table: "occCharge",
          path: "conditionalFormats.add[0].columns",
          column: "fee",
          oldName: "amount",
        },
      ],
    }).join("\n");
    expect(text).toContain(
      "REFUSED 1 · occCharge › conditionalFormats.add[0].columns\n  fee is the name this run's rename gives amount, and alterations.ts names a column by its old name until the run lands.\n  Use the column's old name, amount, this run, then rerun.",
    );
  });

  it("says a new tab's before names no tab, naming its path and the name", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "noSuchBeforeTab",
          table: "Payments",
          path: "newTabs[0].before",
          before: "ledger",
        },
      ],
    }).join("\n");
    expect(text).toContain(
      "REFUSED 1 · Payments › newTabs[0].before\n  newTabs[0].before names ledger, which is neither a tab at the last sync nor a newTabs title.\n  Name a tab the ID-name map holds or a newTabs title, then rerun.",
    );
  });

  it("says a misplaced @ isn't a brace, with the four forms", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "malformedBrace",
          table: "occCharge",
          path: "amount.formula",
          brace: "${@vendor[cost]}",
        },
      ],
    }).join("\n");
    expect(text).toContain("REFUSED 1 · occCharge › amount.formula");
    expect(text).toContain(
      "${@vendor[cost]} isn't a brace: write ${@column} or ${table[@column]} for this row, ${column} or ${table[column]} for the whole column.",
    );
  });

  it("says a NUMBER_* rule's value isn't a number, naming the value and the type", () => {
    const text = render({
      ...appDryRun,
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "notANumber",
          table: "occCharge",
          path: "conditionalFormats.add[0].condition.values[0]",
          conditionType: "NUMBER_EQ",
          value: "FALSE",
        },
      ],
    }).join("\n");
    expect(text).toContain(
      "REFUSED 1 · occCharge › conditionalFormats.add[0].condition.values[0]\n  FALSE isn't a number, which a NUMBER_EQ condition takes: Sheets refuses the whole batch over it.\n  Write a number there, then rerun.",
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

describe("SyncRunReport's head-row overlap stop", () => {
  const stopped = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      { kind: "headRows", table: "vendor", tab: "Vendors", neighbour: "Notes" },
    ],
  });
  const stoppedText = stopped.join("\n");

  it("names the granted Table and the Table its head rows would sit on", () => {
    expect(resultLine(stopped)).toBe(
      "STOPPED: 1 Table can't take head rows. Nothing was sent, and nothing will be until it is resolved.",
    );
    expect(stoppedText).toContain(
      `STOP 1 · vendor › letApiAccess\n  vendor's head rows, the 3 rows above its header on its tab, Vendors, would sit on Table "Notes".\n  Move vendor or "Notes" so that vendor's head rows sit clear of it, then rerun.\n  The entry stays in alterations.ts.`,
    );
  });

  it("counts a grant whose head rows sit on two Tables as one Table", () => {
    const twoNeighbours = render({
      ...appDryRun,
      outcome: { kind: "stopped" },
      stops: [
        { kind: "headRows", table: "vendor", tab: "Vendors", neighbour: "A" },
        { kind: "headRows", table: "vendor", tab: "Vendors", neighbour: "B" },
      ],
    });

    expect(resultLine(twoNeighbours)).toBe(
      "STOPPED: 1 Table can't take head rows. Nothing was sent, and nothing will be until they are resolved.",
    );
  });
});

describe("SyncRunReport's header-taken stop", () => {
  const stopped = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      { kind: "headRows", table: "vendor", tab: "Vendors", headerRow: 6 },
      {
        kind: "headerTaken",
        table: "occCharge",
        header: "Late fee",
        column: "lateFee",
        columnId: "c:och:Q8x1mZp",
      },
    ],
  });
  const stoppedText = stopped.join("\n");

  it("names the header, the existing column and its ID", () => {
    expect(stoppedText).toContain(
      'STOP 2 · occCharge › newColumns\n  occCharge already has a column headed "Late fee": lateFee, c:och:Q8x1mZp.',
    );
  });

  it("counts it as a stop beside a head-row stop", () => {
    expect(resultLine(stopped)).toBe(
      "STOPPED: 1 Table can't take head rows and 1 new column taken. Nothing was sent, and nothing will be until they are resolved.",
    );
    expect(stoppedText).toContain("Resolve each stop above, then rerun");
  });
});

describe("SyncRunReport's before-first-column stop", () => {
  const stopped = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      {
        kind: "beforeFirstColumn",
        table: "occCharge",
        path: "newColumns[1].before",
        header: "Ref",
        before: "id",
      },
    ],
  });
  const stoppedText = stopped.join("\n");

  it("names the entry and words it as the insert's guard does", () => {
    expect(stoppedText).toContain(
      "STOP 1 · occCharge › newColumns[1].before\n  Table occCharge can't take \"Ref\" before id, its first column: a column inserted there lands outside the Table.\n  Name a later column in before, or leave before out to add it at the Table's end.\n  The entry stays in alterations.ts.",
    );
  });

  it("counts it in the result line", () => {
    expect(resultLine(stopped)).toBe(
      "STOPPED: 1 new column before a first column. Nothing was sent, and nothing will be until it is resolved.",
    );
  });
});

describe("SyncRunReport's new-Table and new-tab stops", () => {
  const stopped = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      {
        kind: "tableNameTaken",
        name: "Occ Charge",
        existing: "Occ charge",
        tableId: "tbl-0a1b2c3d4e",
      },
      {
        kind: "tabTitleTaken",
        title: "Charges",
        existing: "charges",
        gid: 4217,
      },
      {
        kind: "tabTitleRepeated",
        title: "Ledger",
        paths: ["newTabs[0]", "newTabs[2]"],
      },
    ],
  });
  const stoppedText = stopped.join("\n");

  it("names the repeated tab title and each entry giving it", () => {
    expect(stoppedText).toContain(
      'STOP 3 · new tab "Ledger"\n  newTabs[0], newTabs[2] give one title\'s codebase name. Give each new tab its own title, or merge the entries.\n  The entries stay in alterations.ts.',
    );
  });

  it("names the new Table, the live one and its ID", () => {
    expect(stoppedText).toContain(
      'STOP 1 · new Table "Occ Charge"\n  A Table already gives its name\'s codebase name: "Occ charge", tbl-0a1b2c3d4e.',
    );
  });

  it("names the new tab, the live one and its GID", () => {
    expect(stoppedText).toContain(
      'STOP 2 · new tab "Charges"\n  A tab already gives its title\'s codebase name: "charges", GID 4217.',
    );
  });

  it("counts each kind of stop", () => {
    expect(resultLine(stopped)).toBe(
      "STOPPED: 1 new Table name taken and 1 new tab title taken and 1 new tab title repeated. Nothing was sent, and nothing will be until they are resolved.",
    );
  });
});

describe("SyncRunReport's created Table", () => {
  const text = render({
    ...appSend,
    outcome: { kind: "sent", summary: "2 alterations landed" },
    landed: [
      {
        kind: "send",
        subject: { tab: "Payments", path: "newTabs" },
        what: "create the tab",
      },
      {
        kind: "createdTable",
        table: "paymentLog",
        tab: "Payments",
        tableId: "tbl-0a1b2c3d4e",
        columns: [
          { header: "Payee", column: "payee", columnId: "c:plg:Xw3nB7d" },
          { header: "Amount", column: "amount", columnId: "c:plg:Q8x1mZp" },
        ],
      },
    ],
  }).join("\n");

  it("prints the Table's codebase name and ID and each column's", () => {
    expect(text).toContain(
      'Payments: created Table paymentLog, tbl-0a1b2c3d4e; created "Payee" as payee, c:plg:Xw3nB7d; created "Amount" as amount, c:plg:Q8x1mZp.',
    );
  });
});

describe("SyncRunReport's no-such-rule stop", () => {
  const ruleHandle = "occCharge·all·=${void}=TRUE·bg#eeeeee";
  const lines = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      {
        kind: "noSuchRule",
        table: "occCharge",
        path: `conditionalFormats.remove["${ruleHandle}"]`,
        handles: [
          "occCharge·all·=${void}=TRUE·bg#f4cccc",
          "occCharge·amount·NUMBER_LESS(0)·text#cc0000",
        ],
      },
    ],
  });
  const text = lines.join("\n");

  it("renders under Stops, with the configs still refreshed", () => {
    expect(sections(lines)).toEqual(["Stops", "Configs", "Next"]);
  });

  it("names the entry and lists the Table's current rule handles", () => {
    expect(text).toContain(
      `STOP 1 · occCharge › conditionalFormats.remove["${ruleHandle}"]\n  no such rule: adopt by removing the entry.`,
    );
    expect(text).toContain(
      "  occCharge's rules now, by handle:\n    occCharge·all·=${void}=TRUE·bg#f4cccc\n    occCharge·amount·NUMBER_LESS(0)·text#cc0000\n  The entry stays in alterations.ts.",
    );
  });
});

describe("SyncRunReport's unmodelable-rule stop", () => {
  const ruleHandle = 'A5:A9 | unmodelable {"gradientRule":{}}';
  const lines = render({
    ...appDryRun,
    outcome: { kind: "stopped" },
    stops: [
      {
        kind: "unmodelableRule",
        table: "occCharge",
        path: `conditionalFormats.replace["${ruleHandle}"]`,
        handle: ruleHandle,
      },
    ],
  });
  const text = lines.join("\n");

  it("names the handle and says the rule can only be removed", () => {
    expect(text).toContain(
      `STOP 1 · occCharge › conditionalFormats.replace["${ruleHandle}"]\n  The rule ${ruleHandle} is unmodelable: sync can't write a rule like it, so it can only be removed.`,
    );
    expect(resultLine(lines)).toBe(
      "STOPPED: 1 unmodelable rule to replace. Nothing was sent, and nothing will be until it is resolved.",
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
      'occCharge: created "Late fee" as lateFee, c:och:Q8x1mZp.',
    );
    expect(text).toContain("occCharge › notes: the formula alteration.");
    expect(text).not.toContain("alterations.ts is now empty.");
  });

  it("prints no held list and no token, only the rerun", () => {
    expect(text).not.toContain("Token");
    expect(text).toContain("Rerun the dry run: npm run app:sync");
  });
});

describe("SyncRunReport, a send Google refused partway", () => {
  const googleRefusal =
    'Moving columns in Table "occCharge" on sheet "Charges" would split Table "beside" beside it. Move "beside" so that it sits entirely within rows 1–4, or entirely outside them.';
  const lines = render({
    ...appSend,
    outcome: { kind: "stoppedPartway", landedFlushes: 1, googleRefusal },
    landed: [],
    notSent: [],
    configs: noConfigChanges,
  });

  it("says the earlier flush landed and Google applied none of the refused one", () => {
    expect(resultLine(lines)).toBe(
      `STOPPED PARTWAY: flush 1 landed; flush 2 was refused: Google refused its batch and applied none of it. ${googleRefusal} Nothing after it was sent.`,
    );
  });
});

describe("SyncRunReport, a send whose fill failed after its first flush landed", () => {
  const lines = render({
    ...appSend,
    outcome: {
      kind: "stoppedPartway",
      landedFlushes: 1,
      fillFailure: "socket hang up",
    },
    landed: [],
    notSent: [],
    configs: noConfigChanges,
  });

  it("words the fill's send failing, quoting the message, and says the configs were still written", () => {
    expect(resultLine(lines)).toBe(
      'STOPPED PARTWAY: flush 1 landed; flush 2, the Fill Row IDs send, failed: "socket hang up". The blank IDs stay blank; the configs were still written.',
    );
    expect(lines.join("\n")).toContain(
      "Then rerun the dry run to fill the blank IDs: npm run app:sync",
    );
  });
});

describe("SyncRunReport, a send Google refused at its first flush", () => {
  const googleRefusal = 'Table "beside" would be split.';
  const lines = render({
    ...appSend,
    outcome: { kind: "stoppedBeforeSending", googleRefusal },
    landed: [],
    notSent: [],
    configs: noConfigChanges,
  });

  it("says nothing was sent", () => {
    expect(resultLine(lines)).toBe(
      `STOPPED: the send was refused before any flush landed: Google refused its batch and applied none of it. ${googleRefusal} Nothing was sent.`,
    );
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

describe("SyncRunReport, a Table a person deleted", () => {
  const lines = render({
    ...appSend,
    outcome: { kind: "sent", summary: "nothing to send" },
    notices: [{ kind: "tableGone", table: "occPayment" }],
  });

  it("names the Table under Notices, saying the configs dropped it", () => {
    expect(sections(lines)).toContain("Notices");
    expect(lines).toContain(
      "occPayment: a person deleted this Table from the sheet, so sync dropped it from the configs and retired its ID prefix.",
    );
  });
});

describe("SyncRunReport, a dry run holding back a rule change", () => {
  const lines = render({
    ...appDryRun,
    outcome: { kind: "ready" },
    notices: [
      {
        kind: "ruleHeldBack",
        subject: { table: "occPayment", path: "conditionalFormats.add[0]" },
        shiftedBy: "occCharge",
      },
    ],
  });

  it("names the rule change by its path under Notices, and the Table whose change shifts it", () => {
    expect(sections(lines)).toContain("Notices");
    expect(lines).toContain(
      "occPayment › conditionalFormats.add[0]: held back: occCharge's column change shifts occPayment's columns this run, so it is sent next run.",
    );
  });
});

describe("SyncRunReport, an alteration already met", () => {
  const alreadyMet = [
    {
      subject: { table: "unit", column: "rent", path: "header" },
      what: "already Monthly rent",
    },
  ];
  const sentLanded = [
    {
      kind: "send" as const,
      subject: { table: "unit", column: "notes", path: "header" },
      what: "set to Remarks",
    },
  ];

  it("lists it on its own line, apart from the alterations sent, on a send", () => {
    const lines = render({
      ...appSend,
      outcome: {
        kind: "sent",
        summary: "1 alteration landed; 1 alteration already met",
      },
      landed: sentLanded,
      alreadyMet,
      hasEmptiedAlterations: true,
    });
    expect(sections(lines)).toEqual([
      "Landed",
      "Already met",
      "Configs",
      "Next",
    ]);
    expect(lines.join("\n")).toContain(
      [
        "== Landed: removed from alterations.ts ==",
        "unit › notes › header: set to Remarks",
        "",
        "== Already met: removed from alterations.ts ==",
        "unit › rent › header: already Monthly rent",
        "alterations.ts is now empty.",
      ].join("\n"),
    );
  });

  it("says --send removes it, on a dry run", () => {
    const lines = render({
      ...appDryRun,
      outcome: { kind: "ready" },
      alreadyMet,
    });
    expect(lines).toContain(
      "== Already met: --send removes from alterations.ts ==",
    );
  });

  it("keeps it in alterations.ts on a stopped send", () => {
    const lines = render({
      ...appSend,
      outcome: { kind: "stopped" },
      alreadyMet,
    });
    expect(lines).toContain("== Already met: still in alterations.ts ==");
    expect(lines).not.toContain("alterations.ts is now empty.");
  });
});

describe("SyncRunReport, an entry the planner skipped", () => {
  const landed = [
    {
      kind: "send" as const,
      subject: { table: "unit", column: "notes", path: "header" },
      what: "set to Remarks",
    },
  ];

  // Sides it can't compare, or a value only the sheet holds: the entry stays in the file yet lists in no section.
  it("doesn't say alterations.ts is empty while the skipped entry is left in it", () => {
    const lines = render({
      ...appSend,
      outcome: { kind: "sent", summary: "1 alteration landed" },
      landed,
      hasEmptiedAlterations: false,
    });
    expect(lines).toContain("unit › notes › header: set to Remarks");
    expect(lines).not.toContain("alterations.ts is now empty.");
  });

  it("doesn't say it when a skipped entry is the only one, beside an already-met one", () => {
    const lines = render({
      ...appSend,
      outcome: {
        kind: "sent",
        summary: "nothing to send; 1 alteration already met",
      },
      alreadyMet: [
        {
          subject: { table: "unit", column: "rent", path: "header" },
          what: "already Monthly rent",
        },
      ],
      hasEmptiedAlterations: false,
    });
    expect(lines).toContain("== Already met: removed from alterations.ts ==");
    expect(lines).not.toContain("alterations.ts is now empty.");
  });

  it("says it once the file holds no entries", () => {
    const lines = render({
      ...appSend,
      outcome: { kind: "sent", summary: "1 alteration landed" },
      landed,
      hasEmptiedAlterations: true,
    });
    expect(lines.join("\n")).toContain(
      "unit › notes › header: set to Remarks\nalterations.ts is now empty.",
    );
  });
});

describe("SyncRunReport, a tab a person deleted", () => {
  const lines = render({
    ...appSend,
    outcome: { kind: "sent", summary: "nothing to send" },
    notices: [
      { kind: "tabGone", tab: "payments" },
      { kind: "tableGone", table: "occPayment" },
    ],
  });

  it("names the tab under Notices, beside its Table's line", () => {
    const tabLine =
      "payments: a person deleted this tab from the sheet, so sync dropped it from the configs.";
    expect(lines).toContain(tabLine);
    expect(lines[lines.indexOf(tabLine) + 1]).toBe(
      "occPayment: a person deleted this Table from the sheet, so sync dropped it from the configs and retired its ID prefix.",
    );
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
      },
    ],
    hasEmptiedAlterations: true,
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
      'occPayment: created "Receipt no." as receiptNo, c:opy:Xw3nB7d.\nalterations.ts is now empty.',
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
