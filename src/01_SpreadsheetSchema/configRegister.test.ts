import { describe, expect, it, vi } from "vitest";

import {
  assertNotType,
  assertType,
  type IsExactly,
} from "../testSupport/typeAssertions";
import {
  type Configs,
  type ConfigSetBase,
  type ConfigsNotRegistered,
  type ConfigsOf,
  installConfigs,
  installedConfigs,
} from "./configRegister";
import type { ColumnConfigsGeneric, makeColumnConfigs } from "./makeConfigs";

type ColumnConfigsWithValueName<VN extends string> = ReturnType<
  typeof makeColumnConfigs<
    VN,
    {
      sheet: {
        column: {
          columnId: string;
          header: string;
          valueName: VN;
          isFormula: boolean;
          emptyValueAllowed: boolean;
          customDefaultValue: null;
        };
      };
    }
  >
>;

describe("Register", () => {
  it("resolves an unaugmented Register to an error type, never ConfigSetBase", () => {
    // eslint-disable-next-line @typescript-eslint/no-empty-object-type -- an unaugmented Register
    type Unaugmented = ConfigsOf<{}>;
    assertType<IsExactly<Unaugmented, ConfigsNotRegistered>>(true);
    assertNotType<IsExactly<Unaugmented, ConfigSetBase>>(false);
    assertNotType<Unaugmented extends ConfigSetBase ? true : false>(false);
  });
  it("resolves this program's augmented Register to its own config set", () => {
    assertNotType<IsExactly<Configs, ConfigsNotRegistered>>(false);
    assertNotType<IsExactly<Configs, ConfigSetBase>>(false);
    assertType<Configs extends ConfigSetBase ? true : false>(true);
  });
  // makeColumnConfigs can't name ValueName without a cycle, so this is where a bad valueName fails.
  it("checks every generated valueName against ValueName once the register resolves", () => {
    assertType<
      Configs["columnConfigs"] extends ColumnConfigsGeneric ? true : false
    >(true);
    type Known = ColumnConfigsWithValueName<"date">;
    type Unknown = ColumnConfigsWithValueName<"notAValueName">;
    assertType<Known extends ColumnConfigsGeneric ? true : false>(true);
    assertNotType<Unknown extends ColumnConfigsGeneric ? true : false>(false);
  });
});

describe("installedConfigs", () => {
  it("throws until the entry call installs configs", async () => {
    vi.resetModules();
    const fresh = await import("./configRegister");
    expect(() => fresh.installedConfigs()).toThrow(
      "Configs have not been installed",
    );
  });
  it("refuses a second, different set once one is installed", () => {
    expect(() => installConfigs(installedConfigs())).not.toThrow();
    expect(() => installConfigs({ ...installedConfigs() })).toThrow(
      "A different set of configs is already installed",
    );
  });
});
