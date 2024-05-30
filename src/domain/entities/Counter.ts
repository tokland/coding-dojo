import { Struct } from "./generic/Struct";

export class Counter extends Struct<{ value: number }>() {
    increment(): Counter {
        return this.add(1);
    }

    decrement(): Counter {
        return this.add(-1);
    }

    private add(value: number): Counter {
        return this._update({ value: this.value + value });
    }
}
