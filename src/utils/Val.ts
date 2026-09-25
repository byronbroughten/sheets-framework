import { SerialDate } from "./SerialDate";

class ValidationError extends Error {}

export function validationError(
  value: unknown,
  notAWhat: string,
): ValidationError {
  return new ValidationError(`value "${value}" is not a ${notAWhat}`);
}

interface PrimitiveValueNamesToTypes {
  string: string;
  number: number;
  boolean: boolean;
  date: SerialDate;
}
export type PrimitiveValueName = keyof PrimitiveValueNamesToTypes;
export type PureValue<VN extends PrimitiveValueName> =
  PrimitiveValueNamesToTypes[VN];

const _isS = {
  string(value: unknown): value is string {
    return typeof value === "string";
  },
  emptyString(value: unknown): value is "" {
    return value === "";
  },
  formula(value: unknown): value is string {
    return typeof value === "string" && value[0] === "=";
  },
  number(value: unknown): value is number {
    return typeof value === "number";
  },
  boolean(value: unknown): value is boolean {
    return typeof value === "boolean";
  },
  date(value: unknown): value is SerialDate {
    return SerialDate.isSerial(value);
  },
};

const _validateS = {
  string: (value: unknown): string => {
    if (_isS.string(value)) {
      return value;
    } else {
      throw validationError(value, "string");
    }
  },
  number: (value: unknown): number => {
    if (_isS.number(value)) {
      return value;
    } else {
      throw validationError(value, "number");
    }
  },
  numberOrEmpty: (value: unknown): number | "" => {
    if (_isS.number(value) || _isS.emptyString(value)) {
      return value;
    } else {
      throw validationError(value, "number or empty string");
    }
  },
  boolean: (value: unknown): boolean => {
    if (_isS.boolean(value)) {
      return value;
    } else {
      throw validationError(value, "boolean");
    }
  },
  date: (value: unknown): SerialDate => {
    if (_isS.date(value)) {
      return value;
    } else {
      throw validationError(value, "date");
    }
  },
};

function assert<T>(
  value: T | null | undefined,
  whatNotFound: string = "Value",
): T {
  if (value === null || value === undefined) {
    throw new Error(`${whatNotFound} not found.`);
  }
  return value;
}

export const Val = {
  is: _isS,
  validate: _validateS,
  assert,
};
