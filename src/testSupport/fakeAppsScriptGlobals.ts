import { vi } from "vitest";

/** Stubs the `Logger` global as a no-op spy, so production `Logger.log(...)` calls don't crash under Node. */
export function stubLogger(): { log: ReturnType<typeof vi.fn> } {
  const logger = { log: vi.fn() };
  vi.stubGlobal("Logger", logger);
  return logger;
}

export type FakeTriggerKind =
  "onEdit" | "onChange" | "monthDay" | "everyMinutes";

export interface FakeTrigger {
  handlerFunction: string;
  kind: FakeTriggerKind;
  detail?: number;
}

export interface FakeToast {
  message: string;
  title: string;
  timeoutSeconds: number;
}

/**
 * Stubs `ScriptApp` and `SpreadsheetApp` with just enough of a fluent trigger
 * builder to cover AppsScript.trigger's usage. Created triggers are tracked
 * in the returned array so tests can assert on what was scheduled/deleted,
 * and every toast (message, title and timeout) shown on the active spreadsheet in `toasts`.
 * A null `spreadsheetId` stands in for a standalone script, which has no active spreadsheet.
 */
export function stubScriptAndSpreadsheetApp({
  spreadsheetId = "fake-bound-spreadsheet",
}: { spreadsheetId?: string | null } = {}): {
  triggers: FakeTrigger[];
  toasts: FakeToast[];
} {
  const triggers: FakeTrigger[] = [];
  const toasts: FakeToast[] = [];

  function record(trigger: FakeTrigger): FakeTrigger {
    triggers.push(trigger);
    return trigger;
  }

  function newTrigger(handlerFunction: string): {
    forSpreadsheet: (_spreadsheet: unknown) => {
      onEdit: () => { create: () => FakeTrigger };
      onChange: () => { create: () => FakeTrigger };
    };
    timeBased: () => {
      onMonthDay: (day: number) => { create: () => FakeTrigger };
      everyMinutes: (minutes: number) => { create: () => FakeTrigger };
    };
  } {
    return {
      forSpreadsheet: (_spreadsheet: unknown) => ({
        onEdit: () => ({
          create: () => record({ handlerFunction, kind: "onEdit" }),
        }),
        onChange: () => ({
          create: () => record({ handlerFunction, kind: "onChange" }),
        }),
      }),
      timeBased: () => ({
        onMonthDay: (day: number) => ({
          create: () =>
            record({ handlerFunction, kind: "monthDay", detail: day }),
        }),
        everyMinutes: (minutes: number) => ({
          create: () =>
            record({
              handlerFunction,
              kind: "everyMinutes",
              detail: minutes,
            }),
        }),
      }),
    };
  }

  vi.stubGlobal("ScriptApp", {
    newTrigger,
    getProjectTriggers: () => [...triggers],
    deleteTrigger: (trigger: FakeTrigger) => {
      const index = triggers.indexOf(trigger);
      if (index !== -1) triggers.splice(index, 1);
    },
  });
  vi.stubGlobal("SpreadsheetApp", {
    getActive: () => {
      if (spreadsheetId === null) return null;
      return {
        getId: () => spreadsheetId,
        toast: (message: string, title: string, timeoutSeconds: number) => {
          toasts.push({ message, title, timeoutSeconds });
        },
      };
    },
  });

  return { triggers, toasts };
}
