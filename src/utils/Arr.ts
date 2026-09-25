import { Val } from "./Val";
export type StrictExtract<T, K extends T> = Extract<T, K>;

export interface IndexRange {
  startIndex: number;
  endIndex: number;
}

export const Arr = {
  indexesFromUntil(from: number, until: number): number[] {
    const indexes: number[] = [];
    for (let i = from; i < until; i++) {
      indexes.push(i);
    }
    return indexes;
  },
  // Ascending indexes collapsed into half-open ranges, so a run costs one request.
  contiguousRanges(indexes: number[]): IndexRange[] {
    return [...indexes]
      .sort((a, b) => a - b)
      .reduce((ranges: IndexRange[], index) => {
        const last = ranges[ranges.length - 1];
        if (last && last.endIndex === index) {
          last.endIndex = index + 1;
        } else {
          ranges.push({ startIndex: index, endIndex: index + 1 });
        }
        return ranges;
      }, []);
  },
  hasDuplicates(arr: unknown[]): boolean {
    return new Set(arr).size !== arr.length;
  },
  compareForSort(a: unknown, b: unknown): number {
    if (typeof a === "number" && typeof b === "number") {
      return a - b;
    }
    if (typeof a === "string" && typeof b === "string") {
      return a.localeCompare(b);
    }
    if (a instanceof Date && b instanceof Date) {
      return a.getTime() - b.getTime();
    }
    const stringA = String(a);
    const stringB = String(b);
    return stringA.localeCompare(stringB);
  },
  sortAscending<A>(arr: A[]): A[] {
    return [...arr].sort((a, b) => {
      return this.compareForSort(a, b);
    });
  },
  sortDescending<A>(arr: A[]): A[] {
    return [...arr].sort((a, b) => {
      return this.compareForSort(b, a);
    });
  },
  oneOrThrow<V>(arr: readonly V[]): V {
    if (arr.length !== 1) {
      throw new Error("There is more than one item in this array.");
    } else return Val.assert(arr[0], "The only item");
  },
  firstOrThrow<V>(arr: readonly V[]): V {
    if (arr.length < 1) {
      throw new Error("This array is empty.");
    } else return Val.assert(arr[0], "The first item");
  },
  lastOrThrow<V>(arr: readonly V[]): V {
    const idx = this.lastIdx(arr);
    if (idx < 0) {
      throw new Error("This array has no last value—it has no value.");
    } else return Val.assert(arr[idx], "The last item");
  },
  getOnlyItem<T>(arr: T[], arrayOf?: string): T {
    const strArrayOf = arrayOf ?? "items";
    if (arr.length < 1) {
      throw new ValueNotFoundError(`The array does not have any ${strArrayOf}`);
    } else if (arr.length > 1) {
      throw new Error(`The array has too many ${strArrayOf}`);
    } else {
      return Val.assert(arr[0], "The only item");
    }
  },
  insert<V>(arr: readonly V[], value: V, idx: number): V[] {
    const nextArr = [...arr];
    nextArr.splice(idx, 0, value);
    return nextArr;
  },
  nextRotatingValue<T>(arr: readonly T[], currentValue: T): T {
    if (arr.length === 0) {
      throw new Error("Cannot get next rotating value of an empty array.");
    }
    const currentIdx = arr.indexOf(currentValue);
    const nextIdx = (currentIdx + 1) % arr.length;
    return Val.assert(arr[nextIdx], "The next rotating value");
  },
  replaceAtIdx<V>(arr: readonly V[], value: V, idx: number): V[] {
    const nextArr = [...arr];
    nextArr[idx] = value;
    return nextArr;
  },
  rmFirstMatchOrThrow<T>(arr: T[], value: T): T[] {
    const index = arr.indexOf(value);
    if (index < 0) {
      throw new ValueNotFoundError(`No value in the array matches "${value}".`);
    }
    const nextArr = [...arr];
    nextArr.splice(index, 1);
    return nextArr;
  },
  rmFirstMatchFastMUTATE(arr: unknown[], value: unknown): void {
    const index = arr.indexOf(value);
    arr.splice(index, 1);
  },
  rmAtIndex<T>(arr: readonly T[], idx: number): T[] {
    this.validateIdxOrThrow(arr, idx);
    const nextArr = [...arr];
    nextArr.splice(idx, 1);
    return nextArr;
  },
  replaceValue<T>(arr: T[], value: T, nextValue: T): T[] {
    const nextArr = [...arr];
    while (true) {
      const index = arr.indexOf(value);
      if (index === -1) break;
      nextArr[index] = nextValue;
    }
    return nextArr;
  },
  upOneDimension<T>(arr: T[], innerArrsLength: number): T[][] {
    return arr.reduce(
      (arrOfArrs, item) => {
        if (arrOfArrs.length > 0) {
          const lastRow = this.lastOrThrow(arrOfArrs);
          if (lastRow.length === innerArrsLength) arrOfArrs.push([item]);
          else lastRow.push(item);
        }
        return arrOfArrs;
      },
      [[]] as T[][],
    );
  },
  indicesOf(arr: readonly unknown[], value: unknown): number[] {
    return arr.flatMap((item, idx) => (item === value ? [idx] : []));
  },
  lastIdx(arr: readonly unknown[]): number {
    return arr.length - 1;
  },
  isLastIdx(arr: readonly unknown[], idx: number): boolean {
    return this.lastIdx(arr) === idx;
  },

  includes<T, U extends T>(arr: readonly U[], elem: T): elem is U {
    return (arr as readonly T[]).includes(elem);
  },
  numsInOffsetLength(offset: number, length: number) {
    return Array.from({ length }, (_, k) => k + offset);
  },
  findAndRmFirst<T>(
    arr: T[],
    fn: (value: T) => boolean,
    mustFind: boolean = false,
  ): T[] {
    const nextArr = [...arr];
    const idx = arr.findIndex(fn);
    if (mustFind && idx === -1) {
      throw new ValueNotFoundError("Value not found to remove.");
    }
    if (idx !== -1) arr.splice(idx, 1);
    return nextArr;
  },
  removeLast<T>(arr: T[]): T[] {
    const nextArr = [...arr];
    nextArr.pop();
    return nextArr;
  },
  findAll<T>(arr: readonly T[], fn: (value: T) => boolean): T[] {
    const workingArr = [...arr];
    const all: T[] = [];
    while (true) {
      const idx = workingArr.findIndex(fn);
      if (idx < 0) return all;
      all.push(Val.assert(workingArr[idx], "The found item"));
      workingArr.splice(idx, 1);
    }
  },
  has<T>(arr: T[], fn: (value: T) => boolean): boolean {
    const value = arr.find(fn);
    if (value === undefined) return false;
    else return true;
  },
  exclude<A, B>(a: readonly A[], b: readonly B[]): Exclude<A, B>[] {
    return a.filter(
      (str) => !(b as readonly unknown[]).includes(str),
    ) as Exclude<A, B>[];
  },
  excludeStrict<A, B extends A>(
    a: readonly A[],
    ...b: readonly B[]
  ): Exclude<A, B>[] {
    return a.filter((str) => !(b as readonly A[]).includes(str)) as Exclude<
      A,
      B
    >[];
  },
  extractStrict<A, B extends A>(
    a: readonly A[],
    ...b: readonly B[]
  ): Extract<A, B>[] {
    return a.filter((str) => b.includes(str as B)) as Extract<A, B>[];
  },
  extractOrder<A, B extends A>(
    // is this useful?
    a: readonly A[],
    b: readonly B[],
  ): Extract<A, B>[] {
    return b.filter((str) => a.includes(str)) as Extract<A, B>[];
  },
  extract<A, B>(a: readonly A[], b: readonly B[]): Extract<A, B>[] {
    return a.filter((str) =>
      (b as readonly unknown[]).includes(str),
    ) as Extract<A, B>[];
  },
  idxOrThrow<T>(arr: readonly T[], finder: (val: T) => boolean): number {
    const idx = arr.findIndex(finder);
    if (idx < 0) {
      throw new ValueNotFoundError("Value not found at any index.");
    }
    return idx;
  },
  validateIdxOrThrow(arr: readonly unknown[], idx: number): true {
    const highestIdx = arr.length - 1;
    if (idx > highestIdx) {
      throw new ValueNotFoundError(
        `The passed array does not have a value at passed idx ${idx}`,
      );
    }
    return true;
  },
  combineWithoutIdenticals<A, B>(a: A[], b: B[]): (A | B)[] {
    return [...new Set([...a, ...b])];
  },
} as const;

class ValueNotFoundError extends Error {}
