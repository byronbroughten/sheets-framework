import type { NotEmpty } from "./cellValues";

export interface ValueSchemaBase<VL = unknown> {
  type: VL;
  makeDefault: MakeDefaultValueBase<VL>;
  strictValidate: ValidateValueBase<VL>;
  // What a blank cell reads as, or null where a blank stays a blank.
  blankReadsAs: NotEmpty<VL> | null;
}

export type ValueSchemaKey = keyof ValueSchemaBase;

type MakeDefaultValueBase<VL> = () => VL;
type ValidateValueBase<VL> = (value: unknown) => NotEmpty<VL>;

export function vsc<VL>(props: ValueSchemaBase<VL>): ValueSchemaBase<VL> {
  return props;
}
