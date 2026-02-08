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
            const tables = workbook.tables;
            const names = workbook.names;
            const activeSheet = workbook.worksheets.getActiveWorksheet();

            // Load properties on collections and objects
            sheets.load("items/name, items/visibility");
            tables.load("items/name");
            names.load("items/name, items/formula");
            activeSheet.load("name");

            await context.sync();

            const metadata: WorkbookMetadata = {
                sheets: [],
                tables: [],
                namedRanges: [],
                activeSheet: activeSheet.name
            };

            // Safely iterate sheets
            if (sheets.items) {
                for (let i = 0; i < sheets.items.length; i++) {
                    const s = sheets.items[i];
                    metadata.sheets.push({
                        name: s.name,
                        rowCount: 0,
                        columnCount: 0,
                        visibility: s.visibility
                    });
                }
            }

            // Safely iterate tables
            if (tables.items) {
                for (let i = 0; i < tables.items.length; i++) {
                    metadata.tables.push({
                        name: tables.items[i].name,
                        sheet: "Unknown",
                        range: "Unknown"
                    });
                }
            }

            // Safely iterate named ranges
            if (names.items) {
                for (let i = 0; i < names.items.length; i++) {
                    metadata.namedRanges.push({
                        name: names.items[i].name,
                        formula: names.items[i].formula
                    });
                }
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
