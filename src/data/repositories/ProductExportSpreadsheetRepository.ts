import _c from "../../domain/entities/generic/Collection";
import { Product } from "../../domain/entities/Product";
import { ProductExportRepository } from "../../domain/entities/ProductExportRepository";
import { Maybe } from "../../utils/ts-utils";
import { Sheet, createWorkbook, sheet, types, workbook } from "../Spreadsheets";

export class ProductExportSpreadsheetRepository implements ProductExportRepository {
    async export(name: string, products: Product[]): Promise<void> {
        const [activeProducts, inactiveProducts] = this.splitProducts(products);

        const productsWorkbook = workbook({
            name: name,
            sheets: [
                this.getProductsSheet("Active Products", activeProducts),
                this.getProductsSheet("Inactive Products", inactiveProducts),
                this.getSummarySheet(products),
            ],
        });

        createWorkbook(productsWorkbook);
    }

    private getProductsSheet(
        name: string,
        products: Product[]
    ): Sheet<{ id: string; title: string; quantity: number; status: string }> {
        const productRowsSortedByTitle = _c(products)
            .uniqWith((product1, product2) => product1.equals(product2))
            .sortBy(product => product.title)
            .map(product => ({ ...product, quantity: product.quantity.value }))
            .value();

        return sheet({
            name: name,
            columns: {
                id: { header: "ID", type: types.string },
                title: { header: "Title", type: types.string },
                quantity: { header: "Quantity", type: types.number },
                status: { header: "Status", type: types.string },
            },
            rows: productRowsSortedByTitle,
        });
    }

    private getSummarySheet(products: Product[]): Sheet<{
        productsCount: number;
        itemsCount: number;
        itemsActiveCount: number;
        itemsInactiveCount: number;
    }> {
        const [activeProducts, inactiveProducts] = this.splitProducts(products);

        return sheet({
            name: "Summary",
            columns: {
                productsCount: { header: "# Products", type: types.number },
                itemsCount: { header: "# Items total", type: types.number },
                itemsActiveCount: { header: "# Items active", type: types.number },
                itemsInactiveCount: { header: "# Items inactive", type: types.number },
            },
            rows: [
                {
                    productsCount: cellNumber(products.length),
                    itemsCount: sumQuantities(products),
                    itemsActiveCount: sumQuantities(activeProducts),
                    itemsInactiveCount: sumQuantities(inactiveProducts),
                },
            ],
        });
    }

    private splitProducts(products: Product[]) {
        const activeProducts = products.filter(product => product.status === "active");
        const inactiveProducts = products.filter(product => product.status === "inactive");
        return [activeProducts, inactiveProducts] as const;
    }
}

function cellNumber(n: number): Maybe<number> {
    return n === 0 ? undefined : n;
}

function sumQuantities(products: Product[]): Maybe<number> {
    return cellNumber(
        _c(products)
            .map(product => product.quantity.value)
            .sum()
    );
}
