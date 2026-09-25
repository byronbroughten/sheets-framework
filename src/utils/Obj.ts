import { merge } from "./Obj/merge";
import { spread } from "./Obj/spread";
import { type RemoveFirstN, Str, type TextJoin } from "./Str";
import { type PrimitiveValueName, type PureValue, Val } from "./Val";

export type StrictOmit<T, K extends keyof T> = Pick<T, Exclude<keyof T, K>>;
export type DistributiveOmit<T, K extends keyof T> = T extends unknown
  ? Omit<T, K>
  : never;

export type StrictPick<T, K extends keyof T> = Pick<T, K>;
export type StrictPickPartial<T, K extends keyof T> = Partial<Pick<T, K>>;
export type StrictOmitPartial<T, K extends keyof T> = Partial<Omit<T, K>>;
export type PickStartsWith<T extends object, S extends string> = {
  [K in keyof T as K extends `${S}${string}` ? K : never]: T[K];
};

type UnionToIntersection<U> = (
  U extends unknown ? (k: U) => void : never
) extends (k: infer I) => void
  ? I
  : never;

// Each entry carries its own outer/inner key, so indexing the flat map by a
// generic key resolves them without a side lookup that would degrade to `any`.
export type FlattenTwoLevels<
  T extends Record<string, Record<string, object>>,
  D extends string,
  ON extends PropertyKey,
  IN extends PropertyKey,
> = UnionToIntersection<
  {
    [K1 in keyof T]: {
      [K2 in keyof T[K1] as TextJoin<K1 & string, K2 & string, D>]: {
        [P in keyof T[K1][K2] | ON | IN]: P extends ON
          ? K1
          : P extends IN
            ? K2
            : T[K1][K2][P & keyof T[K1][K2]];
      };
    };
  }[keyof T]
>;

export type InvertObj<O extends Record<string | number, string | number>> = {
  [K in O[keyof O]]: keyof O;
};

type Keys<T> = [keyof T];
type Values<T> = [T[keyof T]];
type Entries<O extends object> = { [K in keyof O]: [K, O[K]] }[keyof O][];

export type Full<O extends object> = {
  [K in keyof O]-?: O[K];
};

type FlagKeys<Obj, Condition> = {
  [Key in keyof Obj]: Obj[Key] extends Condition ? Key : never;
};
type AllowedKeys<Obj, Condition> = FlagKeys<Obj, Condition>[keyof Obj];
export type SubType<Base, Condition> = Pick<Base, AllowedKeys<Base, Condition>>;

export type PropKeyOfValue<
  O extends object,
  V extends O[keyof O],
> = keyof SubType<O, V>;

export type RemoveFirstNFromKeys<T extends object, N extends number> = {
  [K in keyof T & string as RemoveFirstN<K, N>]: T[K];
};

export type KeyedMap<
  T extends Record<PropertyKey, object>, // was Record<PropertyKey, unknown>
  F extends keyof T[keyof T],
  N extends PropertyKey = "name",
> = Map<T[keyof T][F], { [K in keyof T]: T[K] & { [P in N]: K } }[keyof T]>;

function toKeyedMap<
  const T extends Record<PropertyKey, object>, // was Record<PropertyKey, unknown>
  F extends keyof T[keyof T],
  N extends PropertyKey = "name",
>(obj: T, idField: F, nameField: N = "name" as N): KeyedMap<T, F, N> {
  const map = new Map() as KeyedMap<T, F, N>;

  (Object.keys(obj) as (keyof T)[]).forEach((outerKey) => {
    const entry = obj[outerKey] as Record<PropertyKey, unknown>;
    map.set(
      entry[idField as PropertyKey] as T[keyof T][F],
      {
        ...entry,
        [nameField]: outerKey,
      } as unknown as { [K in keyof T]: T[K] & { [P in N]: K } }[keyof T],
    );
  });

  return map;
}
export const Obj = {
  toKeyedMap,
  pushByKey<
    O extends Record<string, unknown[]>,
    K extends keyof O,
    V extends O[K][number],
  >(obj: O, key: K, value: V) {
    if (!obj[key]) {
      obj[key] = [] as unknown as O[K];
    }
    obj[key].push(value);
  },
  keyByValue<O extends object, V extends O[keyof O]>(
    obj: O,
    value: V,
  ): keyof O {
    const key = Obj.keys(obj).find((key) => obj[key] === value);
    if (!key) {
      throw new Error("Value not found in object");
    }
    return key as keyof O;
  },
  isEmpty(obj: object): boolean {
    return Object.keys(obj).length === 0;
  },
  isKey<O extends Record<string, unknown>>(
    obj: O,
    value: unknown,
  ): value is keyof O {
    return Object.keys(obj).includes(value as string);
  },
  stringifyEqual(a: unknown, b: unknown): boolean {
    return JSON.stringify(a) === JSON.stringify(b);
  },
  isObjToAny(value: unknown): value is object {
    if (value && typeof value === "object") return true;
    else return false;
  },
  isObjToRecord(value: unknown): value is Record<string, unknown> {
    if (value && typeof value === "object") return true;
    else return false;
  },
  pick<O extends object, KS extends keyof O>(obj: O, keys: KS[]): Pick<O, KS> {
    return keys.reduce(
      (objNext, key) => {
        if (key in obj) {
          objNext[key] = obj[key];
        }
        return objNext;
      },
      {} as Pick<O, KS>,
    );
  },
  validatePick<
    O extends object,
    VN extends PrimitiveValueName,
    KS extends keyof O,
  >(obj: O, valueName: VN, ...keys: KS[]): Record<KS, PureValue<VN>> {
    return keys.reduce(
      (objNext, key) => {
        const value = Val.validate[valueName](obj[key]) as PureValue<VN>;
        objNext[key] = value;
        return objNext;
      },
      {} as Record<KS, PureValue<VN>>,
    );
  },
  strictPick<O extends object, KS extends keyof O>(
    obj: O,
    keys: KS[],
  ): StrictPick<O, KS> {
    return keys.reduce(
      (objNext, key) => {
        if (key in obj) {
          objNext[key] = obj[key];
        } else {
          throw new Error(`Key ${String(key)} not in object`);
        }
        return objNext;
      },
      {} as StrictPick<O, KS>,
    );
  },
  pickStartsWith<T extends object, S extends string>(
    obj: T,
    prefix: S,
  ): PickStartsWith<T, S> {
    const result = {} as PickStartsWith<T, S>;
    // for…in types the key as keyof T & string, which startsWith needs; Obj.keys gives keyof T.
    for (const key in obj) {
      if (key.startsWith(prefix)) {
        result[key as unknown as keyof PickStartsWith<T, S>] = obj[
          key
        ] as unknown as PickStartsWith<T, S>[keyof PickStartsWith<T, S>];
      }
    }
    return result;
  },
  removeFirstNFromKeys<T extends object, N extends number>(
    obj: T,
    n: N,
  ): RemoveFirstNFromKeys<T, N> {
    const result = {} as RemoveFirstNFromKeys<T, N>;
    // for…in types the key as keyof T & string, which removeFirstN needs; Obj.keys gives keyof T.
    for (const key in obj) {
      const newKey = Str.removeFirstN(
        key,
        n,
      ) as unknown as keyof RemoveFirstNFromKeys<T, N>;
      result[newKey] = obj[key] as unknown as RemoveFirstNFromKeys<
        T,
        N
      >[typeof newKey];
    }
    return result;
  },
  strictOmit<O extends object, KS extends keyof O>(
    obj: O,
    ...keysToOmit: KS[]
  ): StrictOmit<O, KS> {
    return Obj.keys(obj).reduce(
      (objNext, key) => {
        if (!keysToOmit.includes(key as KS)) {
          (objNext as O)[key] = obj[key];
        }
        return objNext;
      },
      {} as StrictOmit<O, KS>,
    );
  },
  keys<O extends object>(obj: O): Keys<O> {
    return Object.keys(obj) as unknown as Keys<O>;
  },
  keysDepreciated<O extends object>(obj: O): (keyof O & string)[] {
    return Object.keys(obj) as (keyof O & string)[];
  },
  values<T extends object>(t: T): Values<Full<T>> {
    return Object.values(t) as unknown as Values<Full<T>>;
  },
  entries<O extends object>(obj: O): Entries<Full<O>> {
    return Object.entries(obj) as unknown as Entries<Full<O>>;
  },
  mapValues<O extends object, R>(
    obj: O,
    fn: (value: O[keyof O], key: keyof O) => R,
  ): { [K in keyof O]: R } {
    return Obj.keys(obj).reduce(
      (mapped, key) => {
        mapped[key] = fn(obj[key], key);
        return mapped;
      },
      {} as { [K in keyof O]: R },
    );
  },
  propKeysOfValue<O extends object, V extends O[keyof O]>(
    obj: O,
    value: V,
  ): PropKeyOfValue<O, V>[] {
    const keys = Obj.keys(obj).filter((key) => obj[key] === value);
    return keys as PropKeyOfValue<O, V>[];
  },
  invert<O extends Record<string | number, string | number>>(
    obj: O,
  ): { [K in O[keyof O]]: keyof O } {
    const objNext = {} as { [K in O[keyof O]]: keyof O };
    Obj.keys(obj).forEach((key) => {
      objNext[obj[key]] = key;
    });
    return objNext;
  },
  merge,
  spread,
} as const;
