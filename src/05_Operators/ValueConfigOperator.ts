import type { CellValue } from "../00_Source/CellValues/cellValues";
import type { FrameworkValueName } from "../00_Source/CellValues/frameworkValueSchemas";
import {
  type FixedListMember,
  makeImportLine,
  type ValueConfigBase,
  type ValueConfigsBase,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import { tableLayout } from "../01_SpreadsheetSchema/tableLayout";
import type { ColumnRaw } from "../02_SpreadsheetRaw/ColumnRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigOperator } from "./ColumnConfigOperator";
import {
  type ConfigSyncState,
  type SourceColumn,
  sourceLabel,
  SpreadsheetBaseOperator,
  type ValueSource,
} from "./SpreadsheetBaseOperator";

export class ValueConfigOperator extends SpreadsheetBaseOperator {
  static init(): ValueConfigOperator {
    return new ValueConfigOperator(SpreadsheetBaseOperator.initOperatorProps());
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get columnConfigOperator(): ColumnConfigOperator {
    return new ColumnConfigOperator(this.operatorProps);
  }
  get valueConfigSync(): ConfigSyncState["valueConfigSync"] {
    return this.configSyncState.valueConfigSync;
  }
  get valueSources(): ValueSource[] {
    return this.valueConfigSync.valueSources;
  }
  fetchAfterColumnConfigSynced(): void {
    this.columnConfigOperator.assertSyncedToSpreadsheet();
    this.valueConfigSync.valueSources =
      this.columnConfigOperator.valueSources();
    this.valueSources
      .filter((source) => !isLiveReference(source))
      .forEach((source) => {
        this._rawColumn(source).gatherFetchFull();
      });
    this.ss.fetchAllPrepped({ skipFetchingProperties: true });
  }
  newValueConfigs(): ValueConfigsBase {
    return this.valueSources.reduce<Record<string, ValueConfigBase>>(
      (acc, source) => {
        acc[source.valueName] = isLiveReference(source)
          ? { kind: "liveReference" }
          : this._fixedListMembers(source);
        return acc;
      },
      {},
    );
  }
  toFileSource(makeConfigsImport: string): string {
    return [
      `${makeImportLine("makeValueConfigs", makeConfigsImport)}`,
      ``,
      `export const valueConfigs = makeValueConfigs(${JSON.stringify(
        this.newValueConfigs(),
        null,
        2,
      )});`,
      ``,
    ].join("\n");
  }
  private _rawColumn({ tableId, header }: SourceColumn): ColumnRaw {
    return this.ss.raw.table(tableId).columnByHeader(header);
  }
  private _fixedListMembers(source: ValueSource): FixedListMember[] {
    const column = this._rawColumn(source);
    const declaredValueTitle = column.profile.declaredValueTitle();
    const members = column.valueArrFilterEmpty.map((value) => {
      if (declaredValueTitle === undefined) return value;
      return typedMember({
        value,
        memberType: memberTypes[declaredValueTitle],
        label: sourceLabel({
          tableName: column.table.name,
          header: source.header,
        }),
      });
    });
    return [...new Set(members)];
  }
}

// A row added to the source Table is an option at once, so its rows aren't listed.
function isLiveReference({ header }: SourceColumn): boolean {
  return header === tableLayout.nameHeader || header === tableLayout.idHeader;
}

type MemberType = "string" | "number" | "boolean";

const memberTypes: Record<FrameworkValueName, MemberType> = {
  id: "string",
  string: "string",
  number: "number",
  date: "number",
  checkbox: "boolean",
  boolean: "boolean",
};

function typedMember({
  value,
  memberType,
  label,
}: {
  value: CellValue;
  memberType: MemberType;
  label: string;
}): FixedListMember {
  if (memberType === "string") return String(value);
  if (typeof value !== memberType) {
    throw new Error(
      `Value source ${label} declares ${memberType} members, but holds ${JSON.stringify(value)}.`,
    );
  }
  return value;
}
