// PROTOTYPE, throwaway: the cases rendered by render.cjs, for sheets-framework#216.
// Table names and the `opy` prefix are real; columns, counts, IDs and values are illustrative.

function ts(body) {
  return ["```ts", "export const alterations = makeAlterations({", body, "});", "```"].join("\n");
}

const caseA = {
  id: "A",
  task: "Make occPayment's Amount column Currency.",
  context: ["You added this to `src/alterations.ts`:", "", ts(`  tables: {\n    occPayment: { columns: { amount: { columnType: "CURRENCY" } } },\n  },`), "", "Then you ran the dry run."].join("\n"),
  steps: [{
    output: {
      command: "npm run app:sync",
      target: "app",
      mode: "dry run",
      result: { kind: "stopped", why: "1 conflict", until: "it is resolved" },
      stops: [{
        kind: "alteration",
        where: "occPayment › amount › columnType",
        label: "both changed",
        yours: "CURRENCY",
        agreed: "AUTOMATIC",
        live: "NUMBER",
        cause: "A person changed this on the sheet after the last sync, and your entry changes it too.",
      }],
      plan: {
        title: "Plan: what --send would send once nothing stops it",
        items: [{
          what: "occPayment: set amount's type to CURRENCY (resends all 15 columns).",
          notes: ["Resets the style and colours of 4 dropdowns: payerCategory, nonOccupantPayerId, formOfPayment, paymentStatus (dropdownStyleReset: \"allow\")."],
        }],
      },
      configs: ["Refreshed. Held back while conflict 1 is open: occPayment › amount › columnType stays AUTOMATIC in the configs, so the regen doesn't take up the sheet's NUMBER."],
      next: ["Resolve conflict 1, then rerun npm run app:sync."],
    },
  }],
};

const caseB = {
  id: "B",
  task: "In occPayment, make Notes say \"Processed\" whenever Processed is ticked.",
  context: ["You added this to `src/alterations.ts`:", "", ts(`  tables: {\n    occPayment: {\n      columns: { notes: { formula: '=IF(SINGLE(occPayment[Processed]),"Processed","")' } },\n    },\n  },`), "", "Then you ran the dry run."].join("\n"),
  steps: [{
    output: {
      command: "npm run app:sync",
      target: "app",
      mode: "dry run",
      result: { kind: "readyHeld", n: 41 },
      held: {
        count: 41,
        groups: [{
          where: "occPayment › notes",
          what: "the formula alteration overwrites entered data in 41 cells:",
          lines: [
            "opy-0007  \"Paid by cheque #1182\"  → formula",
            "opy-0012  \"Late, see email 3/4\"  → formula",
            "opy-0019  \"Split with Household\"  → formula",
            "… and 38 more in notes (over 20 items, the rest are counted, not listed)",
          ],
        }],
        fullList: ".probe/sync-held.json",
        token: "7f3a9c",
      },
      plan: {
        items: [{
          what: "occPayment › notes: write the formula into all 212 body rows. New rows copy it.",
          notes: ["Overwrites 41 cells of entered data (held, above); the other 171 are blank."],
        }],
      },
      configs: ["Refreshed. No changes."],
      next: ["--send refuses this plan without --confirm 7f3a9c, because of the held items."],
    },
  }],
};

const caseC = {
  id: "C",
  task: "On dev, make item's Weight column a Number column, and colour negative weights red.",
  context: ["You added this to `dev/alterations.ts`:", "", ts(`  tables: {\n    item: {\n      columns: { weight: { columnType: "NUMBER" } },\n      conditionalFormats: {\n        add: [{ columns: ["weight"], condition: "NUMBER_LESS", values: ["0"], format: { textColor: "#cc0000" } }],\n      },\n    },\n  },`), "", "Then you ran the dry run."].join("\n"),
  steps: [{
    output: {
      command: "npm run dev:sync",
      target: "dev",
      mode: "dry run",
      result: { kind: "readyHeld", n: 1 },
      held: {
        count: 1,
        groups: [{
          where: "item",
          what: "dropdown style reset (dropdownStyleReset: \"ask\"):",
          lines: [
            "Sending item's column list resets these dropdowns to plain chips and clears their option colours: status, category.",
            "A quirk of Google's API: any column-list send resets them, and the API can neither read nor set them.",
          ],
        }],
        token: "b21e07",
      },
      plan: {
        items: [
          { what: "item: set weight's type to NUMBER (resends all 9 columns).", notes: ["Resets 2 dropdowns' style and colours (held, above)."] },
          {
            what: "item: add rule item·weight·NUMBER_LESS(0)·text#cc0000.",
            notes: ["Added first; it outranks 1 rule on these cells: item·all·=${status}=\"Retired\"·bg#eeeeee."],
          },
        ],
      },
      notices: ["item: ID is the first column. Sheets' Table name chip covers the cell above it, so item's Fill Row IDs checkbox can't be clicked. Moving ID out of first place fixes it."],
      configs: ["Refreshed. No changes."],
      next: ["--send refuses this plan without --confirm b21e07, because of the held item."],
    },
  }],
};

const caseD = {
  id: "D",
  task: "Delete the occChargeArchive Table and the Old Imports tab; nothing uses them any more.",
  context: ["You added this to `src/alterations.ts`:", "", ts(`  tables: {\n    occChargeArchive: { deleted: true },\n  },\n  tabs: {\n    oldImports: { deleted: true },\n  },`), "", "Then you ran the dry run."].join("\n"),
  steps: [{
    output: {
      command: "npm run app:sync",
      target: "app",
      mode: "dry run",
      result: { kind: "readyHeld", n: 1630 },
      held: {
        count: 1630,
        groups: [
          { where: "occChargeArchive", what: "deleting the Table deletes entered data in 1,204 cells, over 318 rows:", lines: ["description 318 · amount 318 · chargeDate 318 · unitId 250", "(Formula columns aren't held: total, occupantName.)"] },
          { where: "tab Old Imports", what: "deleting the tab deletes:", lines: ["importRaw (a Table without API access): entered data in 380 cells", "outside any Table: 46 non-blank cells, formulas and values"] },
        ],
        fullList: ".probe/sync-held.json",
        token: "4c9d10",
      },
      plan: {
        items: [
          {
            what: "occChargeArchive: delete the Table, its cells and its head rows. Its tab, Charge Archive, stays, empty.",
            notes: [
              "Removes 2 conditional-format rules wholly over it: occChargeArchive·amount·NUMBER_LESS(0)·text#cc0000, occChargeArchive·all·=${void}=TRUE·bg#eeeeee.",
              "Leaves 1 rule partly over it, on Charge Archive A2:H400 (=$H2>1000); Sheets keeps it over the blank cells.",
              "3 formulas outside this Table read it. After the deletion they read blank cells, with no error:",
              "  Ledger!C2        =SUM(occChargeArchive[Amount])",
              "  Ledger!C3        =COUNTA(occChargeArchive[ID])",
              "  Owner Report!F9  =SUMIFS(occChargeArchive[Amount], occChargeArchive[Unit ID], A9)",
              "Retires the ID prefix ocha.",
            ],
          },
          {
            what: "tab Old Imports (gid 88123): delete the tab, with its Table importRaw.",
            notes: ["1 formula on another tab reads it. After the deletion it shows #REF!:", "  Ledger!H1        ='Old Imports'!B2"],
          },
        ],
      },
      configs: ["Refreshed. No changes."],
      next: ["--send refuses this plan without --confirm 4c9d10, because of the held items."],
    },
  }],
};

const caseE = {
  id: "E",
  task: "Add a \"Late fee\" Currency column to occCharge, right before Rent, and make unit's Move in date required.",
  context: ["Earlier this session you added this to `src/alterations.ts`, taking the names from `src/generated/idNameMap.ts`. People have been editing the sheet since.", "", ts(`  tables: {\n    occCharge: { newColumns: [{ header: "Late fee", columnType: "CURRENCY", before: "rent" }] },\n    unit: { columns: { moveIn: { emptyValueAllowed: false } } },\n  },`), "", "Now you ran the dry run."].join("\n"),
  steps: [{
    output: {
      command: "npm run app:sync",
      target: "app",
      mode: "dry run",
      result: { kind: "ready" },
      plan: {
        items: [
          { what: "occCharge: insert column \"Late fee\" (CURRENCY) before monthlyRent.", notes: ["Rewrites 2 conditional-format rules over occCharge after the insert, since Sheets leaves them wrong."] },
          { what: "unit › moveOut: emptyValueAllowed → false. Changes the configs only; nothing is sent to the sheet." },
        ],
      },
      configs: [
        "Renamed by a person since the last sync; the codebase name follows the header:",
        "  occCharge › rent → monthlyRent      header \"Rent\" → \"Monthly rent\"",
        "Names that now point at a different column than at the last sync. Reread src/generated/idNameMap.ts before writing another entry:",
        "  unit › moveIn     was c:unt:k2PqA9x, now c:unt:Rb7sW0e",
        "  unit › moveOut    was c:unt:Rb7sW0e, now c:unt:k2PqA9x",
        "Rewritten in alterations.ts, so each entry still names the column it named before:",
        "  occCharge.newColumns[0].before    \"rent\" → \"monthlyRent\"",
        "  unit.columns.moveIn → unit.columns.moveOut    (c:unt:k2PqA9x, which a person retitled \"Move out\")",
        "Not rewritten: claims and endpoint code. tsc fails where they say rent. A name that moved still type-checks, so code using unit's moveIn or moveOut needs reading by eye.",
      ],
      next: ["Nothing held: npm run app:sync -- --send sends this plan."],
    },
  }],
};

const caseF = {
  id: "F",
  task: "Give vendor API access and add a \"Tax ID\" Text column to it. On occCharge, change the grey \"void rows\" rule to light red, and add a rule colouring Late fee red when it's over 50.",
  context: ["You added this to `src/alterations.ts`:", "", ts(`  tables: {\n    vendor: { letApiAccess: true, newColumns: [{ header: "Tax ID", columnType: "TEXT" }] },\n    occCharge: {\n      conditionalFormats: {\n        replace: {\n          "occCharge·all·=\${void}=TRUE·bg#eeeeee": { columns: "all", formula: "=\${void}=TRUE", format: { backgroundColor: "#f4cccc" } },\n        },\n        add: [{ columns: ["lateFee"], formula: "=\${latefee}>50", format: { textColor: "#cc0000" } }],\n      },\n    },\n  },`), "", "Then you ran the dry run."].join("\n"),
  steps: [
    {
      output: {
        command: "npm run app:sync",
        target: "app",
        mode: "dry run",
        result: { kind: "refused", n: 3 },
        refusals: [
          { where: "vendor › newColumns", text: "vendor has no API access yet, so sync has nothing to compare its columns against.\nGranting access takes its own run: leave letApiAccess: true alone on vendor and run sync; add the column in a later run,\nonce vendor's columns are in idNameMap.ts." },
          { where: "occCharge › conditionalFormats.replace[\"occCharge·all·=${void}=TRUE·bg#eeeeee\"]", text: "No such rule: no rule on occCharge has this handle now. A person changed the rule after the last sync\n(a handle comes from its columns, formula, condition and format). Adopt by removing the entry;\noccCharge's rules as they are now are in src/generated/tableConfigs.ts, under conditionalFormats." },
          { where: "occCharge › conditionalFormats.add[0].formula", text: "${latefee} names no column of occCharge. Its columns: id, occupancyId, description, amount, lateFee, monthlyRent, void, chargeDate." },
        ],
        next: ["Fix each refused entry, then rerun npm run app:sync."],
      },
    },
    {
      before: "## Later\n\nSuppose you made your changes for the first run and ran the dry run again, and it printed this.",
      output: {
        command: "npm run app:sync",
        target: "app",
        mode: "dry run",
        result: { kind: "stopped", why: "1 Table can't take its head rows", until: "it is moved" },
        stops: [{ where: "vendor › letApiAccess", text: "vendor's header row is on row 6 of its tab, Vendors. A Table needs its header on row 4 to take head rows\n(rows 1 to 3: column IDs, action row, heading). Move vendor so its header sits on row 4, then rerun.\nThe entry stays in alterations.ts." }],
        configs: ["Refreshed. No changes."],
        next: ["Resolve stop 1, then rerun npm run app:sync."],
      },
    },
  ],
  ask: [
    "1. After the first run: say exactly what you change, what you run, and anything you tell the developer (quote it).",
    "2. After the second run: say exactly what you do next.",
    "3. What in either output was unclear, or tempted you toward a step you decided against?",
    "4. Last, separately: was anything about how entries in alterations.ts are named unclear?",
  ].join("\n"),
};

const caseG = {
  id: "G",
  task: "Add a \"Late fee\" Currency column to occCharge, and set occCharge's Notes to the formula the developer gave you.",
  context: [
    "You added this to `src/alterations.ts`:",
    "",
    ts(`  tables: {\n    occCharge: {\n      newColumns: [{ header: "Late fee", columnType: "CURRENCY" }],\n      columns: { notes: { formula: '=IF(SINGLE(occCharge[Late fee])>0,"Late","")' } },\n    },\n  },`),
    "",
    "Your dry run held 12 cells of entered data in occCharge › notes, listed by row ID. You pasted the list to the developer, who replied \"yes, those 12 can go\". Then you ran the send.",
  ].join("\n"),
  steps: [{
    output: {
      command: "npm run app:sync -- --send --confirm 91be4f",
      target: "app",
      mode: "send",
      result: { kind: "stoppedPartway", landed: "flush 1 of 2", refused: "flush 2", why: "the sheet changed after the dry run, so --confirm 91be4f doesn't cover its held items" },
      landed: ["occCharge: created column \"Late fee\" as c:och:Q8x1mZp, CURRENCY, at the Table's end."],
      notSent: ["occCharge › notes: the formula alteration."],
      configs: ["Refreshed from the sheet after flush 1: occCharge now has lateFee."],
      next: ["Rerun the dry run: npm run app:sync"],
    },
  }],
};

const caseH = {
  id: "H",
  task: "Move occPayment's Processed column to just before Notes, and add a \"Receipt no.\" Text column at the end.",
  context: ["You added this to `src/alterations.ts`:", "", ts(`  tables: {\n    occPayment: {\n      columns: { processed: { before: "notes" } },\n      newColumns: [{ header: "Receipt no.", columnType: "TEXT" }],\n    },\n  },`), "", "Your dry run held nothing, so you ran the send."].join("\n"),
  steps: [{
    output: {
      command: "npm run app:sync -- --send",
      target: "app",
      mode: "send",
      result: { kind: "sent", summary: "2 alterations landed; 3 blank IDs filled" },
      landed: [
        "occPayment › processed: moved before notes.",
        "occPayment: created column \"Receipt no.\" as c:opy:Xw3nB7d, TEXT, at the Table's end.",
        "alterations.ts is now empty.",
      ],
      plan: {
        title: "Also sent",
        items: [
          { what: "occPayment: rewrote 3 conditional-format rules after the move, since Sheets leaves them wrong." },
          { what: "occPayment: re-spanned 1 rule that ended at row 127 to the Table's last body row, 212: occPayment·all·=${processed}=FALSE·bg#fff2cc." },
          { what: "household: filled 3 blank IDs: hh-8Kd2Lq0, hh-T41xZpa, hh-c9WmE2r (rows 140 to 142)." },
        ],
      },
      notices: ["occPayment: 1 hand-drawn rule couldn't be rewritten, and the move left it pointing at the wrong column: on tab Occ Payment, A3:N3 (it covers a head row), =$H3=\"\". Fix it by hand in Sheets, or remove it with a conditionalFormats.remove entry."],
      configs: [
        "Typed traits a person changed since the last sync, on Tables with API access. tsc fails where code relied on the old ones:",
        "  household › phone    columnType TEXT → NUMBER",
      ],
      next: ["sync wrote src/generated/ and src/alterations.ts. Run npm run check, then commit them with your change."],
    },
  }],
};

module.exports = [caseA, caseB, caseC, caseD, caseE, caseF, caseG, caseH];
