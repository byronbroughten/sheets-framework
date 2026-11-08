import { describe, expect, it } from "vitest";

import type {
  ClaimedTraits,
  PendingAlterations,
} from "../01_SpreadsheetSchema/configReaders/intendedState";
import {
  SyncPlanner,
  type SyncPlannerProps,
  type TraitReadings,
} from "./SyncPlanner";

const subject = { table: "item", column: "name", path: "header" };
const formulaSubject = {
  table: "item",
  column: "requiredCount",
  path: "formula",
};

function readings(header: string | undefined): TraitReadings {
  return {
    tables: {
      item: {
        traits: {},
        columns: { name: header === undefined ? {} : { header } },
      },
    },
    tabs: {},
  };
}

function formulaReadings(formula: string): TraitReadings {
  return {
    tables: { item: { traits: {}, columns: { requiredCount: { formula } } } },
    tabs: {},
  };
}

const headerAlteration: PendingAlterations = {
  tables: { item: { columns: { name: { header: "Mine" } } } },
};

function planner(props: Partial<SyncPlannerProps>): SyncPlanner {
  return SyncPlanner.init({
    claims: [],
    alterations: {},
    tabsByTable: {},
    forcedTraits: [],
    ...props,
  });
}

function planHeader(
  previous: string | undefined,
  live: string,
  intended = "Mine",
) {
  return planner({
    alterations: {
      tables: { item: { columns: { name: { header: intended } } } },
    },
  }).plan({ previous: readings(previous), live: readings(live) });
}

function liveHeaderOf(configTraits: TraitReadings): unknown {
  return configTraits.tables.item?.columns.name?.header;
}

describe("SyncPlanner.plan's three-way table", () => {
  it.each([
    {
      row: "agreed",
      previous: "Mine",
      live: "Mine",
      sends: 0,
      stops: 0,
      takes: "Mine",
    },
    {
      row: "agent's change",
      previous: "Old",
      live: "Old",
      sends: 1,
      stops: 0,
      takes: "Mine",
    },
    {
      row: "person's change",
      previous: "Mine",
      live: "Theirs",
      sends: 0,
      stops: 1,
      takes: "Mine",
    },
    {
      row: "both landed on one value",
      previous: "Old",
      live: "Mine",
      sends: 0,
      stops: 0,
      takes: "Mine",
    },
    {
      row: "both changed",
      previous: "Old",
      live: "Theirs",
      sends: 0,
      stops: 1,
      takes: "Old",
    },
  ])("$row", ({ previous, live, sends, stops, takes }) => {
    const result = planHeader(previous, live);
    expect(result.plan).toHaveLength(sends);
    expect(result.stops).toHaveLength(stops);
    expect(liveHeaderOf(result.configTraits)).toBe(takes);
  });

  it("sends the intended value under its trait", () => {
    expect(planHeader("Old", "Old").plan).toEqual([
      { subject, what: "set to Mine", effects: [] },
    ]);
  });

  it("names all three values on a stop", () => {
    expect(planHeader("Old", "Theirs").stops).toEqual([
      expect.objectContaining({
        kind: "conflict",
        subject,
        yours: "Mine",
        agreed: "Old",
        live: "Theirs",
      }),
    ]);
  });

  it("leaves a trait not intended to the sheet", () => {
    const previous: TraitReadings = {
      tables: {
        item: {
          traits: {},
          columns: { name: { header: "Old", formula: "=1" } },
        },
      },
      tabs: {},
    };
    const live: TraitReadings = {
      tables: {
        item: {
          traits: {},
          columns: { name: { header: "New", formula: "=1" } },
        },
      },
      tabs: {},
    };
    const result = planner({}).plan({ previous, live });
    expect(result.plan).toEqual([]);
    expect(result.stops).toEqual([]);
    expect(liveHeaderOf(result.configTraits)).toBe("New");
  });

  it("compares a column type by its content, not its key order", () => {
    const columnType = {
      type: "DROPDOWN",
      rule: { condition: "ONE_OF_LIST", values: ["a"] },
    };
    const reordered = {
      rule: { values: ["a"], condition: "ONE_OF_LIST" },
      type: "DROPDOWN",
    };
    const sheet: TraitReadings = {
      tables: {
        item: { traits: {}, columns: { name: { columnType: reordered } } },
      },
      tabs: {},
    };
    const result = planner({
      alterations: { tables: { item: { columns: { name: { columnType } } } } },
    }).plan({ previous: sheet, live: sheet });
    expect(result.plan).toEqual([]);
    expect(result.stops).toEqual([]);
  });

  it("applies no normaliser: a value differing only in case is a change", () => {
    expect(planHeader("Old", "Old", "old").plan).toHaveLength(1);
  });

  it("sends a Table's and a tab's traits too", () => {
    const sheet: TraitReadings = {
      tables: { item: { traits: { name: "item" }, columns: {} } },
      tabs: { log: { title: "Log" } },
    };
    const result = planner({
      alterations: {
        tables: { item: { name: "stock" } },
        tabs: { log: { title: "Ledger" } },
      },
    }).plan({ previous: sheet, live: sheet });
    expect(result.plan.map((item) => item.subject)).toEqual([
      { table: "item", path: "name" },
      { tab: "log", path: "title" },
    ]);
  });
});

describe("a trait the previous configs don't record", () => {
  it("takes live as its baseline, so a matching value is quiet", () => {
    const result = planHeader(undefined, "Mine");
    expect(result.plan).toEqual([]);
    expect(result.stops).toEqual([]);
  });

  it("takes live as its baseline, so a differing value is the agent's change, never a conflict", () => {
    const result = planHeader(undefined, "Theirs");
    expect(result.stops).toEqual([]);
    expect(result.plan).toHaveLength(1);
  });
});

describe("a key in intended state and the configs but not live", () => {
  const claims = [
    {
      declaredBy: "itemEndpoint",
      traits: {
        item: { columns: { requiredCount: { formula: "=1" } } },
      } as ClaimedTraits,
    },
  ];
  const noColumn: TraitReadings = {
    tables: { item: { traits: {}, columns: {} } },
    tabs: {},
  };

  it("stops a claim on a column a person deleted", () => {
    const result = planner({ claims }).plan({
      previous: formulaReadings("=1"),
      live: noColumn,
    });
    expect(result.stops).toEqual([
      expect.objectContaining({
        kind: "conflict",
        subject: formulaSubject,
        claimedBy: "itemEndpoint",
      }),
    ]);
  });

  it("keeps the column's agreed value in the configs, so the next run still stops", () => {
    const result = planner({ claims }).plan({
      previous: formulaReadings("=1"),
      live: noColumn,
    });
    expect(
      result.configTraits.tables.item?.columns.requiredCount?.formula,
    ).toBe("=1");
  });

  it("stays a stop under force, having no column to send to", () => {
    const result = planner({
      claims,
      forcedTraits: [
        { table: "item", column: "requiredCount", trait: "formula" },
      ],
    }).plan({ previous: formulaReadings("=1"), live: noColumn });
    expect(result.stops).toHaveLength(1);
    expect(result.plan).toEqual([]);
    expect(result.forcedTraits).toEqual([]);
  });

  it("isn't a conflict for a Table a person deleted", () => {
    const result = planner({ alterations: headerAlteration }).plan({
      previous: readings("Mine"),
      live: { tables: {}, tabs: {} },
    });
    expect(result.stops).toEqual([]);
    expect(result.plan).toEqual([]);
  });
});

describe("a claim", () => {
  const claims = [
    {
      declaredBy: "itemEndpoint",
      traits: {
        item: { columns: { requiredCount: { formula: "=1" } } },
      } as ClaimedTraits,
    },
  ];

  it("stops on a person's edit, naming its declarer", () => {
    const result = planner({ claims }).plan({
      previous: formulaReadings("=1"),
      live: formulaReadings("=2"),
    });
    expect(result.stops).toEqual([
      expect.objectContaining({
        claimedBy: "itemEndpoint",
        yours: "=1",
        live: "=2",
      }),
    ]);
  });
});

describe("a trait in conflict", () => {
  it("is held back at its last agreed value", () => {
    const result = planHeader("Old", "Theirs");
    expect(result.configs.heldBack).toEqual([{ subject, value: "Old" }]);
    expect(liveHeaderOf(result.configTraits)).toBe("Old");
  });

  it("holds nothing back when nothing conflicts", () => {
    expect(planHeader("Old", "Old").configs.heldBack).toEqual([]);
  });

  it("stops every send, leaving an unrelated change at live in the configs", () => {
    const sheet = (header: string, formula: string): TraitReadings => ({
      tables: {
        item: {
          traits: {},
          columns: { name: { header }, requiredCount: { formula } },
        },
      },
      tabs: {},
    });
    const result = planner({
      alterations: {
        tables: {
          item: {
            columns: {
              name: { header: "Mine" },
              requiredCount: { formula: "=2" },
            },
          },
        },
      },
    }).plan({ previous: sheet("Old", "=1"), live: sheet("Theirs", "=1") });
    expect(result.stops).toHaveLength(1);
    expect(result.plan).toHaveLength(1);
    expect(
      result.configTraits.tables.item?.columns.requiredCount?.formula,
    ).toBe("=1");
  });
});

describe("force", () => {
  const forcedTraits = [{ table: "item", column: "name", trait: "header" }];

  function planForced(previous: string, live: string) {
    return planner({ alterations: headerAlteration, forcedTraits }).plan({
      previous: readings(previous),
      live: readings(live),
    });
  }

  it("turns that one stop into a send and a held item", () => {
    const result = planForced("Old", "Theirs");
    expect(result.stops).toEqual([]);
    expect(result.plan).toEqual([
      { subject, what: "set to Mine, over the sheet's Theirs", effects: [] },
    ]);
    expect(result.forcedTraits).toEqual(forcedTraits);
    expect(result.configs.heldBack).toEqual([]);
    expect(liveHeaderOf(result.configTraits)).toBe("Mine");
  });

  it("adds no held item for a trait that wasn't stopped", () => {
    const result = planForced("Old", "Old");
    expect(result.forcedTraits).toEqual([]);
  });

  it("leaves other stops standing, and holds the forced trait back while they do", () => {
    const sheet = (header: string, formula: string): TraitReadings => ({
      tables: {
        item: {
          traits: {},
          columns: { name: { header }, requiredCount: { formula } },
        },
      },
      tabs: {},
    });
    const result = planner({
      alterations: {
        tables: {
          item: {
            columns: {
              name: { header: "Mine" },
              requiredCount: { formula: "=2" },
            },
          },
        },
      },
      forcedTraits,
    }).plan({ previous: sheet("Old", "=1"), live: sheet("Theirs", "=3") });
    expect(result.stops).toHaveLength(1);
    expect(result.configs.heldBack).toContainEqual({ subject, value: "Old" });
    expect(liveHeaderOf(result.configTraits)).toBe("Old");
  });
});

describe("SyncPlanner.refusals", () => {
  const claimOne = {
    declaredBy: "itemEndpoint",
    traits: {
      item: { columns: { requiredCount: { formula: "=1" } } },
    } as ClaimedTraits,
  };
  const claimTwo = {
    declaredBy: "stockChore",
    traits: {
      item: { columns: { requiredCount: { formula: "=2" } } },
    } as ClaimedTraits,
  };

  it("are empty for declarers that agree", () => {
    const refusals = planner({
      claims: [claimOne, { ...claimOne, declaredBy: "stockChore" }],
      alterations: {
        tables: { item: { columns: { requiredCount: { formula: "=1" } } } },
      },
    }).refusals;
    expect(refusals).toEqual([]);
  });

  it("name two claims that disagree", () => {
    expect(planner({ claims: [claimOne, claimTwo] }).refusals).toEqual([
      {
        kind: "disagreeingDeclarers",
        table: "item",
        path: "requiredCount.formula",
        declarers: [
          { declaredBy: "itemEndpoint", value: "=1" },
          { declaredBy: "stockChore", value: "=2" },
        ],
      },
    ]);
  });

  it("name a claim and an alteration that disagree", () => {
    const refusals = planner({
      claims: [claimOne],
      alterations: {
        tables: { item: { columns: { requiredCount: { formula: "=9" } } } },
      },
    }).refusals;
    expect(refusals).toEqual([
      expect.objectContaining({
        kind: "disagreeingDeclarers",
        declarers: [
          { declaredBy: "itemEndpoint", value: "=1" },
          { declaredBy: "alterations.ts", value: "=9" },
        ],
      }),
    ]);
  });

  function deletionRefusals(props: Partial<SyncPlannerProps>) {
    return planner(props).refusals.flatMap((refusal) =>
      refusal.kind === "touchesDeletion"
        ? [{ deleted: refusal.deleted, touchedBy: refusal.touchedBy }]
        : [],
    );
  }

  it.each<{ row: string; props: Partial<SyncPlannerProps>; touchedBy: string }>(
    [
      {
        row: "a sibling's before",
        props: {
          alterations: {
            tables: {
              item: {
                columns: { name: { deleted: true }, id: { before: "name" } },
              },
            },
          },
        },
        touchedBy: "tables.item.columns.id.before",
      },
      {
        row: "a new column's before",
        props: {
          alterations: {
            tables: {
              item: {
                columns: { name: { deleted: true } },
                newColumns: [{ header: "X", before: "name" }],
              },
            },
          },
        },
        touchedBy: "tables.item.newColumns[0].before",
      },
      {
        row: "a formula's brace in another column",
        props: {
          alterations: {
            tables: {
              item: {
                columns: {
                  name: { deleted: true },
                  optionalNote: { formula: "=${name}" },
                },
              },
            },
          },
        },
        touchedBy: "tables.item.columns.optionalNote.formula",
      },
      {
        row: "a pending rule's columns",
        props: {
          alterations: {
            tables: {
              item: {
                columns: { name: { deleted: true } },
                conditionalFormats: {
                  add: [
                    {
                      columns: ["name"],
                      condition: { type: "NOT_BLANK" },
                      format: { bold: true },
                    },
                  ],
                },
              },
            },
          },
        },
        touchedBy: "tables.item.conditionalFormats.add[0].columns",
      },
      {
        row: "a claim on the column",
        props: {
          claims: [
            {
              declaredBy: "itemEndpoint",
              traits: {
                item: { columns: { name: { formula: "=1" } } },
              } as ClaimedTraits,
            },
          ],
          alterations: {
            tables: { item: { columns: { name: { deleted: true } } } },
          },
        },
        touchedBy: "claim (itemEndpoint)",
      },
    ],
  )("name a column's deletion beside $row", ({ props, touchedBy }) => {
    expect(deletionRefusals(props)).toEqual([
      { deleted: { table: "item", column: "name" }, touchedBy },
    ]);
  });

  it("name a Table's deletion beside a brace in another Table", () => {
    expect(
      deletionRefusals({
        alterations: {
          tables: {
            item: { deleted: true },
            log: { columns: { amount: { formula: "=${item[name]}" } } },
          },
        },
      }),
    ).toEqual([
      {
        deleted: { table: "item" },
        touchedBy: "tables.log.columns.amount.formula",
      },
    ]);
  });

  it("name a tab's deletion beside an entry on a Table it holds", () => {
    expect(
      deletionRefusals({
        tabsByTable: { item: "itemTab" },
        alterations: {
          tables: { item: { name: "stock" } },
          tabs: { itemTab: { deleted: true } } as PendingAlterations["tabs"],
        },
      }),
    ).toEqual([{ deleted: { tab: "itemTab" }, touchedBy: "tables.item" }]);
  });

  it("name a tab's deletion beside another tab's before", () => {
    expect(
      deletionRefusals({
        alterations: {
          tabs: { log: { deleted: true }, item: { before: "log" } },
        },
      }),
    ).toEqual([{ deleted: { tab: "log" }, touchedBy: "tabs.item.before" }]);
  });

  it("let a deletion stand alone", () => {
    expect(
      planner({ alterations: { tables: { item: { deleted: true } } } })
        .refusals,
    ).toEqual([]);
  });
});
