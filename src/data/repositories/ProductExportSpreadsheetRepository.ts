import ExcelJS from "exceljs";
import { Product } from "../../domain/entities/Product";
import { ProductExportRepository } from "../../domain/entities/ProductExportRepository";
import { Maybe } from "../../utils/ts-utils";
import _c from "../../domain/entities/generic/Collection";

type Workbook = {
    name: string;
    sheets: Sheet<string>[];
};

type Sheet<Field extends string> = {
    name: string;
    columns: Record<Field, { header: string }>;
    rows: Array<Record<Field, string | number | undefined>>;
};

export class ProductExportSpreadsheetRepository implements ProductExportRepository {
    async export(name: string, products: Product[]): Promise<void> {
        const [activeProducts, inactiveProducts] = this.splitProducts(products);

        const workbook: Workbook = {
            name: name,
            sheets: [
                this.getProductsSheet("Active Products", activeProducts),
                this.getProductsSheet("Inactive Products", inactiveProducts),
                this.getSummarySheet(products),
            ],
        };

        this.createWorkbook(workbook);
    }

    private splitProducts(products: Product[]): [Product[], Product[]] {
        const activeProducts = products.filter(product => product.status === "active");
        const inactiveProducts = products.filter(product => product.status === "inactive");
        return [activeProducts, inactiveProducts];
    }

    private getProductsSheet(
        name: string,
        products: Product[]
    ): Sheet<"id" | "title" | "quantity" | "status"> {
        const productRowsSortedByTitle = _c(products)
            .uniqWith((product1, product2) => product1.equals(product2))
            .sortBy(product => product.title)
            .map(product => ({ ...product, quantity: product.quantity.value }))
            .value();

        return {
            name: name,
            columns: {
                id: { header: "ID" },
                title: { header: "Title" },
                quantity: { header: "Quantity" },
                status: { header: "Status" },
            },
            rows: productRowsSortedByTitle,
        };
    }

    private getSummarySheet(
        products: Product[]
    ): Sheet<"productsCount" | "itemsCount" | "itemsActiveCount" | "itemsInactiveCount"> {
        const [activeProducts, inactiveProducts] = this.splitProducts(products);

        return {
            name: "Summary",
            columns: {
                productsCount: { header: "# Products" },
                itemsCount: { header: "# Items total" },
                itemsActiveCount: { header: "# Items active" },
                itemsInactiveCount: { header: "# Items inactive" },
            },
            rows: [
                {
                    productsCount: cellNumber(products.length),
                    itemsCount: sumQuantities(products),
                    itemsActiveCount: sumQuantities(activeProducts),
                    itemsInactiveCount: sumQuantities(inactiveProducts),
                },
            ],
        };
    }

    private createWorkbook(workbook: Workbook): Promise<void> {
        const ejsWorkbook = new ExcelJS.Workbook();

        workbook.sheets.forEach(sheet => {
            const ejsSheet = ejsWorkbook.addWorksheet(sheet.name);
            const ejsColumns = Object.entries(sheet.columns).map(([key, column]) => ({
                header: column.header,
                key: key,
            }));

            ejsSheet.columns = ejsColumns;
            ejsSheet.addRows(sheet.rows);
        });

        return ejsWorkbook.xlsx.writeFile(workbook.name);
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
