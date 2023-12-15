import ExcelJS from "exceljs";

// All this section can be moved to another file

/* Workbook data entity */

type BaseSchema = Record<string, string | number>;

export type Workbook = {
    name: string;
    sheets: Sheet<BaseSchema>[];
};

export type Sheet<Schema extends BaseSchema> = {
    name: string;
    columns: Record<keyof Schema, { header: string }>;
    rows: Array<RowDefinition<Schema>>;
};

type RowDefinition<Schema extends BaseSchema> = {
    [K in keyof Schema]: Schema[K] | undefined;
};

type BaseColumns = Record<string, { header: string; type: string | number }>;

export function sheet<Columns extends BaseColumns>(options: {
    name: string;
    columns: Columns;
    rows: Array<{
        [K in keyof Columns]: Columns[K]["type"] | undefined;
    }>;
}): Sheet<{ [K in keyof Columns]: Columns[K]["type"] }> {
    return options;
}

export function workbook(options: Workbook): Workbook {
    return options;
}

export const types = { string: "", number: 0 };

// Workbook writer using ExcelJS

export function createWorkbook(workbook: Workbook): Promise<void> {
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

const _workbookExample = workbook({
    name: "Example",
    sheets: [productsSheet],
});
