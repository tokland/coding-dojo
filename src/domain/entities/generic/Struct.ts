export function Struct<Attrs>() {
    abstract class Base {
        constructor(_attributes: Attrs) {
            Object.assign(this, _attributes);
        }

        _getAttributes(): Attrs {
            const entries = Object.getOwnPropertyNames(this).map(key => [key, (this as any)[key]]);
            return Object.fromEntries(entries) as Attrs;
        }

        protected _update(partialAttrs: Partial<Attrs>): this {
            const ParentClass = this.constructor as new (values: Attrs) => typeof this;
            return new ParentClass({ ...this._getAttributes(), ...partialAttrs });
        }

        static create<U extends Base>(this: new (attrs: Attrs) => U, attrs: Attrs): U {
            return new this(attrs);
        }
    }

    return Base as {
        new (values: Attrs): Attrs & Base;
        create: (typeof Base)["create"];
    };
}

const GenericStruct = Struct<unknown>();

export type GenericStructInstance = InstanceType<typeof GenericStruct>;
