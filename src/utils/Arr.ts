import { Val } from "./Val";

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
  numsInOffsetLength(offset: number, length: number) {
    return Array.from({ length }, (_, k) => k + offset);
  },
  indexesOf(arr: readonly unknown[], value: unknown): number[] {
    return arr.flatMap((item, index) => (item === value ? [index] : []));
  },
  lastIndex(arr: readonly unknown[]): number {
    return arr.length - 1;
  },
  isLastIndex(arr: readonly unknown[], index: number): boolean {
    return Arr.lastIndex(arr) === index;
  },
  indexOrThrow<T>(arr: readonly T[], finder: (val: T) => boolean): number {
    const index = arr.findIndex(finder);
    if (index < 0) {
      throw new Error("Value not found at any index.");
    }
    return index;
  },
  validateIndexOrThrow(arr: readonly unknown[], index: number): true {
    if (index > Arr.lastIndex(arr)) {
      throw new Error(
        `The passed array does not have a value at passed idx ${index}`,
      );
    }
    return true;
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
      return Arr.compareForSort(a, b);
    });
  },
  sortDescending<A>(arr: A[]): A[] {
    return [...arr].sort((a, b) => {
      return Arr.compareForSort(b, a);
    });
  },

  oneOrThrow<V>(arr: readonly V[]): V {
    if (arr.length < 1) {
      throw new Error("This array is empty.");
    }
    if (arr.length > 1) {
      throw new Error("There is more than one item in this array.");
    }
    return Val.assert(arr[0], "The only item");
  },
  firstOrThrow<V>(arr: readonly V[]): V {
    if (arr.length < 1) {
      throw new Error("This array is empty.");
    } else return Val.assert(arr[0], "The first item");
  },
  lastOrThrow<V>(arr: readonly V[]): V {
    const index = Arr.lastIndex(arr);
    if (index < 0) {
      throw new Error("This array has no last value—it has no value.");
    } else return Val.assert(arr[index], "The last item");
  },
  getOnlyItem<T>(arr: T[], arrayOf?: string): T {
    const strArrayOf = arrayOf ?? "items";
    if (arr.length < 1) {
      throw new Error(`The array does not have any ${strArrayOf}`);
    } else if (arr.length > 1) {
      throw new Error(`The array has too many ${strArrayOf}`);
    } else {
      return Val.assert(arr[0], "The only item");
    }
  },
  nextRotatingValue<T>(arr: readonly T[], currentValue: T): T {
    if (arr.length === 0) {
      throw new Error("Cannot get next rotating value of an empty array.");
    }
    const currentIndex = arr.indexOf(currentValue);
    const nextIndex = (currentIndex + 1) % arr.length;
    return Val.assert(arr[nextIndex], "The next rotating value");
  },

  insert<V>(arr: readonly V[], value: V, index: number): V[] {
    const nextArr = [...arr];
    nextArr.splice(index, 0, value);
    return nextArr;
  },
  replaceAtIndex<V>(arr: readonly V[], value: V, index: number): V[] {
    const nextArr = [...arr];
    nextArr[index] = value;
    return nextArr;
  },
  replaceValue<T>(arr: T[], value: T, nextValue: T): T[] {
    return arr.map((item) => (item === value ? nextValue : item));
  },
  upOneDimension<T>(arr: T[], innerArrsLength: number): T[][] {
    return arr.reduce(
      (arrOfArrs, item) => {
        if (arrOfArrs.length > 0) {
          const lastRow = Arr.lastOrThrow(arrOfArrs);
          if (lastRow.length === innerArrsLength) arrOfArrs.push([item]);
          else lastRow.push(item);
        }
        return arrOfArrs;
      },
      [[]] as T[][],
    );
  },

  removeFirstMatchOrThrow<T>(arr: T[], value: T): T[] {
    const index = arr.indexOf(value);
    if (index < 0) {
      throw new Error(`No value in the array matches "${value}".`);
    }
    const nextArr = [...arr];
    nextArr.splice(index, 1);
    return nextArr;
  },
  removeFirstMatchInPlace(arr: unknown[], value: unknown): void {
    const index = arr.indexOf(value);
    arr.splice(index, 1);
  },
  removeAtIndex<T>(arr: readonly T[], index: number): T[] {
    Arr.validateIndexOrThrow(arr, index);
    const nextArr = [...arr];
    nextArr.splice(index, 1);
    return nextArr;
  },
  findAndRemoveFirst<T>(
    arr: T[],
    fn: (value: T) => boolean,
    mustFind: boolean = false,
  ): T[] {
    const nextArr = [...arr];
    const index = nextArr.findIndex(fn);
    if (mustFind && index === -1) {
      throw new Error("Value not found to remove.");
    }
    if (index !== -1) nextArr.splice(index, 1);
    return nextArr;
  },
  removeLast<T>(arr: T[]): T[] {
    const nextArr = [...arr];
    nextArr.pop();
    return nextArr;
  },

  hasDuplicates(arr: unknown[]): boolean {
    return new Set(arr).size !== arr.length;
  },
  includes<T, U extends T>(arr: readonly U[], elem: T): elem is U {
    return (arr as readonly T[]).includes(elem);
  },
  has<T>(arr: T[], fn: (value: T) => boolean): boolean {
    const value = arr.find(fn);
    if (value === undefined) return false;
    else return true;
  },
  findAll<T>(arr: readonly T[], fn: (value: T) => boolean): T[] {
    const workingArr = [...arr];
    const all: T[] = [];
    while (true) {
      const index = workingArr.findIndex(fn);
      if (index < 0) return all;
      all.push(Val.assert(workingArr[index], "The found item"));
      workingArr.splice(index, 1);
    }
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
  extract<A, B>(a: readonly A[], b: readonly B[]): Extract<A, B>[] {
    return a.filter((str) =>
      (b as readonly unknown[]).includes(str),
    ) as Extract<A, B>[];
  },
  extractStrict<A, B extends A>(
    a: readonly A[],
    ...b: readonly B[]
  ): Extract<A, B>[] {
    return a.filter((str) => b.includes(str as B)) as Extract<A, B>[];
  },
  extractOrder<A, B extends A>(a: readonly A[], b: readonly B[]): Extract<A, B>[] {
    return b.filter((str) => a.includes(str)) as Extract<A, B>[];
  },
  combineWithoutIdenticals<A, B>(a: A[], b: B[]): (A | B)[] {
    return [...new Set([...a, ...b])];
  },
} as const;
