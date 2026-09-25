import type {
  UniformRowName,
  UniformRowValue,
  UniformRowValueName,
} from "../../00_Source/CellValues/cellValues";
import { uniformRows } from "../../01_SpreadsheetSchema/uniformRows";
import type { StrictOmit } from "../../utils/Obj";
import type { RowRawProps } from "./RowBaseRaw";
import { RowCommonRaw } from "./RowCommonRaw";

export interface RowUniformProps<UN extends UniformRowName> extends StrictOmit<
  RowRawProps,
  "rowIndex"
> {
  uniformRowName: UN;
}

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
