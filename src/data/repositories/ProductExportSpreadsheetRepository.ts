import ExcelJS from "exceljs";
import { Product } from "../../domain/entities/Product";
import { ProductExportRepository } from "../../domain/entities/ProductExportRepository";
import { Maybe } from "../../utils/ts-utils";
import _c from "../../domain/entities/generic/Collection";

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

        createWorkbook(workbook);
    }

    private splitProducts(products: Product[]) {
        const activeProducts = products.filter(product => product.status === "active");
        const inactiveProducts = products.filter(product => product.status === "inactive");
        return [activeProducts, inactiveProducts] as const;
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

    private getProductsSheet(
        name: string,
        products: Product[]
    ): Sheet<{ id: string; title: string; quantity: number; status: string }> {
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

// All this section can be moved to another file

/* Workbook data entity */

type BaseSchema = Record<string, string | number>;

type Workbook = {
    name: string;
    sheets: Sheet<BaseSchema>[];
};

type Sheet<Schema extends BaseSchema> = {
    name: string;
    columns: Record<keyof Schema, { header: string }>;
    rows: Array<{ [K in keyof Schema]: Schema[K] | undefined }>;
};

// Builder

type BaseColumns = Record<string, { header: string; type: string | number }>;

function sheet<Columns extends BaseColumns>(options: {
    name: string;
    columns: Columns;
    rows: Array<{
        [K in keyof Columns]: Columns[K]["type"] | undefined;
    }>;
}): Sheet<{ [K in keyof Columns]: Columns[K]["type"] }> {
    return options;
}

function workbook(options: Workbook): Workbook {
    return options;
}

const types = { string: "", number: 0 };

// Workbook writer using ExcelJS

function createWorkbook(workbook: Workbook): Promise<void> {
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

/* Example */

const productsSheet = sheet({
    name: "Products",
    columns: {
        id: { header: "ID", type: types.string },
        title: { header: "Title", type: types.string },
        quantity: { header: "Quantity", type: types.number },
        status: { header: "Status", type: types.string },
    },
    rows: [
        { id: "1", title: "Shoes", quantity: 5, status: "enabled" },
        { id: "2", title: "T-Shirt", quantity: 10, status: "disabled" },
    ],
});

const workbookExample = workbook({
    name: "Example",
    sheets: [productsSheet],
});
