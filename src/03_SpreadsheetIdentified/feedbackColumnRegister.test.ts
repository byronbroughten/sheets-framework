import { beforeEach, describe, expect, it, vi } from "vitest";

import type * as RegisterModule from "./feedbackColumnRegister";

type Register = typeof RegisterModule;

async function freshRegister(): Promise<Register> {
  vi.resetModules();
  return import("./feedbackColumnRegister");
}

describe("installFeedbackColumnIds", () => {
  let register: Register;
  beforeEach(async () => {
    register = await freshRegister();
  });

  it("gives an empty set until one is installed", () => {
    expect(register.installedFeedbackColumnIds()).toEqual(new Map());
  });

  it("gives back the installed set", () => {
    register.installFeedbackColumnIds(new Map([["item", new Set(["c:a"])]]));

    expect(register.installedFeedbackColumnIds()).toEqual(
      new Map([["item", new Set(["c:a"])]]),
    );
  });

  it("accepts the same set again, built afresh", () => {
    register.installFeedbackColumnIds(
      new Map([["item", new Set(["c:a", "c:b"])]]),
    );

    expect(() =>
      register.installFeedbackColumnIds(
        new Map([["item", new Set(["c:b", "c:a"])]]),
      ),
    ).not.toThrow();
  });

  it("refuses a different set once one is installed", () => {
    register.installFeedbackColumnIds(new Map([["item", new Set(["c:a"])]]));

    expect(() =>
      register.installFeedbackColumnIds(new Map([["item", new Set(["c:b"])]])),
    ).toThrow("A different set of feedback columns is already installed");
    expect(() =>
      register.installFeedbackColumnIds(
        new Map([
          ["item", new Set(["c:a"])],
          ["runItem", new Set(["c:a"])],
        ]),
      ),
    ).toThrow("A different set of feedback columns is already installed");
  });
});
