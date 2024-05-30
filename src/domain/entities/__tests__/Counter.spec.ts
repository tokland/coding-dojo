import { describe, expect, test } from "vitest";
import { Counter } from "../Counter";

describe("Counter", () => {
    test("increment", () => {
        const counter0 = getZeroCounter();

        const counterIncremented = counter0.increment();

        expect(counterIncremented.value).toEqual(1);
        expect(counter0.value, "counter should be immutable").toEqual(0);
    });

    test("decrement", () => {
        const counter0 = getZeroCounter();

        const counterDecremented = counter0.decrement();

        expect(counterDecremented.value).toEqual(-1);
        expect(counter0.value, "counter should be immutable").toEqual(0);
    });
});

function getZeroCounter() {
    return Counter.create({ value: 0 });
}
