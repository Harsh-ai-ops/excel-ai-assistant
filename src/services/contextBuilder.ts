/* global Excel */

/* global Excel */
export interface WorkbookMetadata {
    sheets: {
        name: string;
        rowCount: number;
        columnCount: number;
        visibility: string;
    }[];
    tables: {
        name: string;
        sheet: string;
        range: string;
    }[];
    namedRanges: {
        name: string;
        formula: string;
    }[];
    activeSheet: string;
}

export class ContextBuilder {
    static async getWorkbookMetadata(): Promise<WorkbookMetadata> {
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const sheets = workbook.worksheets;
            sheets.load("items/name, items/visibility");

            const tables = workbook.tables;
            tables.load("items/name, items/worksheet, items/range");

            const names = workbook.names;
            names.load("items/name, items/formula");

            const activeSheet = workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");

            await context.sync();

            const metadata: WorkbookMetadata = {
                sheets: [],
                tables: [],
                namedRanges: [],
                activeSheet: activeSheet.name
            };

            for (const sheet of sheets.items) {
                // For row/column count, we'd need to load the used range per sheet
                // To keep it fast, we'll just list names for now
                metadata.sheets.push({
                    name: sheet.name,
                    rowCount: 0, // Placeholder
                    columnCount: 0, // Placeholder
                    visibility: sheet.visibility
                });
            }

            for (const table of tables.items) {
                metadata.tables.push({
                    name: table.name,
                    sheet: table.worksheet ? "Unknown" : "Unknown", // Worksheets need separate load usually
                    range: "Unknown"
                });
            }

            for (const name of names.items) {
                metadata.namedRanges.push({
                    name: name.name,
                    formula: name.formula
                });
            }

            return metadata;
        });
    }

    static buildContextString(metadata: WorkbookMetadata, activeSheetData?: string): string {
        let context = `WORKBOOK CONTEXT:\n`;
        context += `Active Sheet: ${metadata.activeSheet}\n`;

        context += `\nSheets:\n`;
        metadata.sheets.forEach(s => {
            context += `- ${s.name} (${s.visibility})\n`;
        });

        if (metadata.tables.length > 0) {
            context += `\nTables:\n`;
            metadata.tables.forEach(t => {
                context += `- ${t.name}\n`;
            });
        }

        if (metadata.namedRanges.length > 0) {
            context += `\nNamed Ranges:\n`;
            metadata.namedRanges.forEach(n => {
                context += `- ${n.name} (${n.formula})\n`;
            });
        }

        if (activeSheetData) {
            context += `\nACTIVE SHEET DATA:\n${activeSheetData}`;
        }

        return context;
    }
}
