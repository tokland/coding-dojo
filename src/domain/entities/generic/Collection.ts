import { HashMap } from "./HashMap";

/**
 * Wrap a collection of values, expanding methods for Javascript Arrays. An example:
 *
 * ```
 * import _ from "./Collection";
 *
 * const values = _(["1", "2", "3", "3", "4"])
 *     .map(x => parseInt(x))
 *     .filter(x => x > 1)
 *     .uniq()
 *     .reverse()
 *     .value(); // [4, 3, 2]
 * ```
 */

export class Collection<T> {
    protected xs: T[];
    protected constructor(values: T[]) {
        this.xs = values;
    }

    /* Builders */

    static from<T>(xs: T[]): Collection<T> {
        return new Collection(xs);
    }

    static range(start: number, end: number, step = 1): Collection<number> {
        const output = [];
        for (let idx = start; idx < end; idx = idx + step) output.push(idx);
        return Collection.from(output);
    }

    /* Unwrappers */

    value(): T[] {
        return this.xs;
    }

    toArray(): T[] {
        return this.value();
    }

    get size() {
        return this.xs.length;
    }

    /* Methods that return a Collection */

    map<U>(fn: (x: T) => U): Collection<U> {
        return _c(this.xs.map(fn));
    }

    isomap(fn: (x: T) => T): Collection<T> {
        return _c(this.xs.map(fn));
    }

    flatten(): Flatten<T> {
        return _c(this.xs.flat()) as Flatten<T>;
    }

    flatMap<U>(fn: (x: T) => Collection<U>): Collection<U> {
        return _c(this.xs.flatMap(x => fn(x).toArray()));
    }

    select(pred: (x: T) => boolean): Collection<T> {
        return _c(this.xs.filter(pred));
    }

    filter(pred: (x: T) => boolean): Collection<T> {
        return this.select(pred);
    }

    reject(pred: (x: T) => boolean): Collection<T> {
        return _c(this.xs.filter(x => !pred(x)));
    }

    enumerate(): Collection<[number, T]> {
        return _c(this.xs.map((x, idx) => [idx, x]));
    }

    compact(): Collection<NonNullable<T>> {
        return this.reject(x => x === undefined || x === null) as unknown as Collection<NonNullable<T>>;
    }

    compactMap<U>(fn: (x: T) => U | undefined | null): Collection<U> {
        return this.map(fn).compact() as unknown as Collection<U>;
    }

    append(x: T): Collection<T> {
        return _c(this.xs.concat([x]));
    }

    includes(x: T): boolean {
        return this.xs.includes(x);
    }

    every(pred: (x: T) => boolean): boolean {
        return this.xs.every(pred);
    }

    all(pred: (x: T) => boolean): boolean {
        return this.every(pred);
    }

    some(pred: (x: T) => boolean): boolean {
        return this.xs.some(pred);
    }

    any(pred: (x: T) => boolean): boolean {
        return this.some(pred);
    }

    find(pred: (x: T) => boolean): T | undefined {
        return this.xs.find(pred);
    }

    sort(): Collection<T> {
        return this.sortWith(defaultCompareFn);
    }

    reverse(): Collection<T> {
        return _c([...this.xs].reverse());
    }

    sortWith(compareFn: CompareFn<T>): Collection<T> {
        return _c(this.xs.slice().sort(compareFn));
    }

    sortBy<U>(fn: (x: T) => U, options: { compareFn?: CompareFn<U> } = {}): Collection<T> {
        const compareFn = options.compareFn || defaultCompareFn;
        // TODO: Schwartzian transform: decorate + sort tuple + undecorate
        return this.sortWith((a, b) => compareFn(fn(a), fn(b)));
    }

    orderBy(items: OrderItem<T>[]): Collection<T> {
        return this.sortWith((a, b) => {
            return compareArray(a, b, items);
        });
    }

    first(): T | undefined {
        return this.xs[0];
    }

    last(): T | undefined {
        return this.xs[this.xs.length - 1];
    }

    sum(): NumberIfNumber<T> {
        return (this.xs as number[]).reduce((acc, x) => acc + x, 0) as NumberIfNumber<T>;
    }

    take(n: number): Collection<T> {
        return _c(this.xs.slice(0, n));
    }

    drop(n: number): Collection<T> {
        return _c(this.xs.slice(n));
    }

    pairwise(): Collection<[T, T]> {
        const n = 2;

        return _c(
            this.xs
                .slice(0, this.xs.length - n + 1)
                .map((_x, idx) => [this.xs[idx], this.xs[idx + 1]] as [T, T])
        );
    }

    prepend(x: T) {
        return _c([x, ...this.xs]);
    }

    tap(fn: (xs: Collection<T>) => void) {
        fn(this);
        return this;
    }

    splitAt(indexes: number[]): Collection<Collection<T>> {
        return _c(indexes)
            .prepend(0)
            .append(this.xs.length)
            .pairwise()
            .map(([i1, i2]) => _c(this.xs.slice(i1, i2)));
    }

    thru<U>(fn: (xs: Collection<T>) => Collection<U>) {
        return fn(this);
    }

    join(char: string): string {
        return this.xs.join(char);
    }

    get(idx: number): T | undefined {
        return this.xs[idx];
    }

    getMany(idxs: number[]): Collection<T | undefined> {
        return _c(idxs.map(idx => this.xs[idx]));
    }

    intersperse(value: T): Collection<T> {
        return this.flatMap(x => _c([x, value])).thru(cs => cs.take(cs.size - 1));
    }

    uniq(): Collection<T> {
        return this.uniqBy(x => x);
    }

    uniqBy<U>(mapper: (value: T) => U): Collection<T> {
        const seen = new Set<U>();
        const output: Array<T> = [];

        for (const item of this.xs) {
            const mapped = mapper(item);
            if (!seen.has(mapped)) {
                output.push(item);
                seen.add(mapped);
            }
        }

        return _c(output);
    }

    reduce<U>(mapper: (acc: U, value: T) => U, initialAcc: U): U {
        return this.xs.reduce(mapper, initialAcc);
    }

    chunk(size: number): Collection<T[]> {
        return Collection.range(0, this.xs.length, size).map(index => {
            return this.xs.slice(index, index + size);
        });
    }

    cartesian(): T extends Array<infer U> ? Collection<U[]> : never {
        const [ys, ...zss] = this.xs;

        if (!ys) {
            return _c([[]]) as any;
        } else {
            return _c(ys as T[]).flatMap(x =>
                _c(zss)
                    .cartesian()
                    .map(zs => [x, ...zs])
            ) as any;
        }
    }

    forEach(fn: (value: T) => void): void {
        this.xs.forEach(fn);
    }

    zipLongest<S>(xs: Collection<S>): Collection<[T | undefined, S | undefined]> {
        const max = Math.max(this.size, xs.size);
        const pairs = Collection.range(0, max)
            .map(i => [this.xs[i], xs.xs[i]] as [T | undefined, S | undefined])
            .value();
        return _c(pairs);
    }

    zip<S>(xs: Collection<S>): Collection<[T, S]> {
        const min = Math.min(this.size, xs.size);
        return Collection.range(0, min).map(i => [this.xs[i], xs.xs[i]] as [T, S]);
    }

    /* Methods that return HashMap */

    indexBy<U>(grouperFn: (x: T) => U): HashMap<U, T> {
        const initialValue = HashMap.empty<U, T>();

        return this.reduce((acc, x) => {
            const key = grouperFn(x);
            return acc.set(key, x);
        }, initialValue);
    }

    keyBy(grouperFn: (x: T) => string): HashMap<string, T> {
        return this.indexBy(grouperFn);
    }

    groupBy<U>(grouperFn: (x: T) => U): HashMap<U, T[]> {
        const map = this.reduce((acc, value) => {
            const key = grouperFn(value);
            const valuesForKey = acc.get(key) || [];
            valuesForKey.push(value);
            return acc.set(key, valuesForKey);
        }, new Map<U, T[]>());

        return HashMap.fromPairs(Array.from(map.entries()));
    }

    groupFromMap<U, W>(pairGrouperFn: (x: T) => [U, W]): HashMap<U, W[]> {
        const map = this.reduce((acc, x) => {
            const [key, value] = pairGrouperFn(x);
            const valuesForKey = acc.get(key) || [];
            valuesForKey.push(value);
            return acc.set(key, valuesForKey);
        }, new Map<U, W[]>());

        return HashMap.fromPairs(Array.from(map.entries()));
    }

    toHashMap<K, V>(toPairFn: (x: T) => [K, V]): HashMap<K, V> {
        const pairs = this.map(toPairFn).toArray();
        return HashMap.fromPairs(pairs);
    }
}

type NumberIfNumber<T> = T extends number ? number : never;

type Flatten<T> = T extends Array<infer U> ? Collection<U> : never;

/* Sorting */

type CompareRes = -1 | 0 | 1;

type CompareFn<T> = (a: T, b: T) => CompareRes;

type Direction = "asc" | "desc";

function defaultCompareFn<T>(a: T, b: T, direction: Direction = "asc"): CompareRes {
    const [value1, value2] = direction === "asc" ? [1 as const, -1 as const] : [-1 as const, 1 as const];
    return a > b ? value1 : b > a ? value2 : 0;
}

function compareArray<T>(a: T, b: T, items: OrderItem<T>[]): CompareRes {
    const item = items[0];
    if (!item) return 0;
    const [mapper, direction] = item;
    const res = defaultCompareFn(mapper(a), mapper(b), direction);
    return res !== 0 ? res : compareArray(a, b, items.slice(1));
}

type OrderItem<T> = [(obj: T) => unknown, "asc" | "desc"];

/* Direct function invocation style */

type CollectionFns = {
    map: <T, U>(xs: T[], fn: (x: T) => U) => U[];
    isomap: <T>(xs: T[], fn: (x: T) => T) => T[];
    flatten: <T>(xs: T[][]) => T[];
    flatMap: <T, U>(xs: T[], fn: (x: T) => U[]) => U[];
    select: <T>(xs: T[], pred: (x: T) => boolean) => T[];
    reject: <T>(xs: T[], pred: (x: T) => boolean) => T[];
    enumerate: <T>(xs: T[]) => [number, T][];
    compact: <T>(xs: (T | null | undefined)[]) => NonNullable<T>[];
    compactMap: <T, U>(xs: T[], fn: (x: T) => U | undefined | null) => NonNullable<U>[];
    append: <T>(xs: T[], x: T) => T[];
    includes: <T>(xs: T[], x: T) => boolean;
    every: <T>(xs: T[], pred: (x: T) => boolean) => boolean;
    some: <T>(xs: T[], pred: (x: T) => boolean) => boolean;
    find: <T>(xs: T[], pred: (x: T) => boolean) => T | undefined;
    sort: <T>(xs: T[]) => T[];
    reverse: <T>(xs: T[]) => T[];
    sortWith: <T>(xs: T[], compareFn: (a: T, b: T) => number) => T[];
    sortBy: <T, U>(xs: T[], fn: (x: T) => U, options: { compareFn?: (a: U, b: U) => number }) => T[];
    orderBy: <T>(xs: T[], items: OrderItem<T>[]) => T[];
    first: <T>(xs: T[]) => T | undefined;
    last: <T>(xs: T[]) => T | undefined;
    sum: <T>(xs: number[]) => NumberIfNumber<T>;
    take: <T>(xs: T[], n: number) => T[];
    drop: <T>(xs: T[], n: number) => T[];
    pairwise: <T>(xs: T[]) => [T, T][];
    prepend: <T>(xs: T[], x: T) => T[];
    tap: <T>(xs: T[], fn: (xs: T[]) => void) => T[];
    splitAt: <T>(xs: T[], indexes: number[]) => T[][];
    thru: <T, U>(xs: T[], fn: (xs: T[]) => U[]) => U[];
    join: <T>(xs: T[], char: string) => string;
    get: <T>(xs: T[], idx: number) => T | undefined;
    getMany: <T>(xs: T[], idxs: number[]) => (T | undefined)[];
    intersperse: <T>(xs: T[], value: T) => T[];
    uniq: <T>(xs: T[]) => T[];
    uniqBy: <T, U>(xs: T[], mapper: (value: T) => U) => T[];
    reduce: <T, U>(xs: T[], mapper: (acc: U, value: T) => U, initialAcc: U) => U;
    chunk: <T>(xs: T[], size: number) => T[][];
    cartesian: <T>(xs: T[][]) => T[][];
    zipLongest: <T, S>(xs: T[], ys: S[]) => [T | undefined, S | undefined][];
    zip: <T, S>(xs: T[], ys: S[]) => [T, S][];
    indexBy: <T, U>(xs: T[], grouperFn: (x: T) => U) => HashMap<U, T>;
    groupBy: <T, U>(xs: T[], grouperFn: (x: T) => U) => HashMap<U, T[]>;
    groupFromMap: <T, U, W>(xs: T[], pairGrouperFn: (x: T) => [U, W]) => HashMap<U, W[]>;
    toHashMap: <T, K, V>(xs: T[], toPairFn: (x: T) => [K, V]) => HashMap<K, V>;
};

function getClassMethods(cls: any) {
    return Object.entries(Object.getOwnPropertyDescriptors(cls.prototype))
        .map(([name, descriptor]) => (typeof descriptor.value === "function" ? name : null))
        .filter(Boolean) as string[];
}

const fns = Object.fromEntries(
    getClassMethods(Collection).map(name => {
        function fn(xs: any, ...args: any[]) {
            const collection = Collection.from(xs);
            const method = (collection as any)[name].bind(collection);
            const result = method(...args);
            return result instanceof Collection ? result.toArray() : result;
        }
        return [name, fn];
    })
) as CollectionFns;

interface CollectionObj extends CollectionFns {
    <T>(xs: T[]): Collection<T>;
}

const _c = (<T>(xs: T[]) => Collection.from(xs)) as CollectionObj;
Object.assign(_c, fns);

export default _c;
