import ExcelJS from "exceljs";

/* Workbook data entity */

export type Workbook = {
    name: string;
    sheets: Sheet<BaseSchema>[];
};

export type Sheet<Schema extends BaseSchema> = {
    name: string;
    columns: Record<keyof Schema, { header: string }>;
    rows: Array<Row<Schema>>;
};

type BaseSchema = Record<string, string | number>;

type Row<Schema extends BaseSchema> = {
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
