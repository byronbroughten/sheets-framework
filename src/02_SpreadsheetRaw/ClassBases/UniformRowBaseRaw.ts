import type {
  UniformRowName,
  UniformRowValue,
  UniformRowValueName,
} from "../../00_Source/CellValues/cellValues";
import { uniformRows } from "../../01_SpreadsheetSchema/uniformRows";
import { RowCommonRaw } from "./RowCommonRaw";
import type { TableRawProps } from "./TableBaseRaw";

export type RowUniformProps<UN extends UniformRowName> = TableRawProps & {
  uniformRowName: UN;
};

export class UniformRowBaseRaw<
  UN extends UniformRowName,
  VN extends UniformRowValueName<UN> = UniformRowValueName<UN>,
> extends RowCommonRaw {
  readonly uniformRowName: UN;
  constructor({ uniformRowName, ...rest }: RowUniformProps<UN>) {
    super({
      ...rest,
      rowIndex: uniformRows.index(uniformRowName),
    });
    this.uniformRowName = uniformRowName;
    this.validateUniformState();
  }
  get valueName(): VN {
    return this.schema.uniformValueName(this.uniformRowName) as VN;
  }
  get activeValueArr(): (UniformRowValue<UN> | "")[] {
    return [...this.rowState.values()].map((cellState) => cellState.value) as (
      UniformRowValue<UN> | ""
    )[];
  }
  validateUniformState(): void {
    this.validateUniformRowIndex();
    this.ensureStateExists();
  }
  private validateUniformRowIndex(): void {
    this.schema.validateUniformRowIndex(this.rowIndex, this.uniformRowName);
  }
}
