/**
 * Excel Service
 * Handles all Excel API interactions using Office.js
 */

export interface SheetData {
    name: string;
    usedRange: string;
    values: any[][];
    formulas: any[][];
    rowCount: number;
    columnCount: number;
}

export interface WorkbookContext {
    activeSheetName: string;
    sheets: string[];
    activeSheetData: SheetData | null;
}

export class ExcelService {
    /**
     * Check if we're running inside Excel
     */
    static isExcelAvailable(): boolean {
        return typeof Excel !== 'undefined' && Office.context.host === Office.HostType.Excel;
    }

    /**
     * Get basic workbook information
     */
    static async getWorkbookContext(): Promise<WorkbookContext> {
        if (!this.isExcelAvailable()) {
            return {
                activeSheetName: 'Demo Sheet',
                sheets: ['Demo Sheet'],
                activeSheetData: null,
            };
        }

        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const sheets = workbook.worksheets;
            const activeSheet = workbook.worksheets.getActiveWorksheet();

            sheets.load('items/name');
            activeSheet.load('name');

            await context.sync();

            const sheetNames = sheets.items.map((s) => s.name);
            const activeSheetName = activeSheet.name;

            // Get active sheet data
            const activeSheetData = await this.getSheetData(activeSheetName);

            return {
                activeSheetName,
                sheets: sheetNames,
                activeSheetData,
            };
        });
    }

    /**
     * Get data from a specific sheet
     */
    static async getSheetData(sheetName: string): Promise<SheetData> {
        if (!this.isExcelAvailable()) {
            return {
                name: sheetName,
                usedRange: 'A1:C3',
                values: [
                    ['Header 1', 'Header 2', 'Header 3'],
                    ['Data 1', 100, 200],
                    ['Data 2', 150, 300],
                ],
                formulas: [
                    ['', '', ''],
                    ['', '', ''],
                    ['', '', ''],
                ],
                rowCount: 3,
                columnCount: 3,
            };
        }

        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const usedRange = sheet.getUsedRange();

            usedRange.load('address, values, formulas, rowCount, columnCount');

            await context.sync();

            return {
                name: sheetName,
                usedRange: usedRange.address,
                values: usedRange.values,
                formulas: usedRange.formulas,
                rowCount: usedRange.rowCount,
                columnCount: usedRange.columnCount,
            };
        });
    }

    /**
     * Get currently selected range data
     */
    static async getSelectedRange(): Promise<{ address: string; values: any[][] }> {
        if (!this.isExcelAvailable()) {
            return {
                address: 'A1:B2',
                values: [['Sample', 100], ['Data', 200]],
            };
        }

        return await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load('address, values');
            await context.sync();

            return {
                address: range.address,
                values: range.values,
            };
        });
    }

    /**
     * Write value to a cell
     */
    static async writeToCell(
        sheetName: string,
        address: string,
        value: any
    ): Promise<void> {
        if (!this.isExcelAvailable()) {
            console.log(`Would write "${value}" to ${sheetName}!${address}`);
            return;
        }

        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const range = sheet.getRange(address);
            range.values = [[value]];
            await context.sync();
        });
    }

    /**
     * Write formula to a cell
     */
    static async writeFormula(
        sheetName: string,
        address: string,
        formula: string
    ): Promise<void> {
        if (!this.isExcelAvailable()) {
            console.log(`Would write formula "${formula}" to ${sheetName}!${address}`);
            return;
        }

        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const range = sheet.getRange(address);
            range.formulas = [[formula]];
            await context.sync();
        });
    }

    /**
     * Navigate to and select a specific cell
     */
    static async navigateToCell(
        sheetName: string,
        address: string
    ): Promise<void> {
        if (!this.isExcelAvailable()) {
            console.log(`Would navigate to ${sheetName}!${address}`);
            return;
        }

        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            sheet.activate();
            const range = sheet.getRange(address);
            range.select();
            await context.sync();
        });
    }

    /**
     * Build a context string for the LLM about the current workbook state
     */
    static async buildContextForLLM(): Promise<string> {
        const workbook = await this.getWorkbookContext();

        let context = `CURRENT EXCEL WORKBOOK STATE:\n`;
        context += `Active Sheet: "${workbook.activeSheetName}"\n`;
        context += `Available Sheets: ${workbook.sheets.join(', ')}\n\n`;

        if (workbook.activeSheetData) {
            const data = workbook.activeSheetData;
            context += `ACTIVE SHEET DATA (${data.usedRange}):\n`;
            context += `Rows: ${data.rowCount}, Columns: ${data.columnCount}\n\n`;

            // Format data as a simple table representation
            if (data.values && data.values.length > 0) {
                // Show first 20 rows max to avoid token limits
                const maxRows = Math.min(data.values.length, 20);
                context += `Data Preview (first ${maxRows} rows):\n`;

                for (let i = 0; i < maxRows; i++) {
                    const row = data.values[i];
                    const rowLabel = i + 1;
                    context += `Row ${rowLabel}: ${row.join(' | ')}\n`;
                }

                if (data.values.length > maxRows) {
                    context += `... and ${data.values.length - maxRows} more rows\n`;
                }
            }

            // Include formulas if present
            const formulaCells: string[] = [];
            data.formulas.forEach((row: any[], rowIndex) => {
                row.forEach((formula, colIndex) => {
                    if (formula && typeof formula === 'string' && formula.startsWith('=')) {
                        const colLetter = String.fromCharCode(65 + colIndex);
                        formulaCells.push(`${colLetter}${rowIndex + 1}: ${formula}`);
                    }
                });
            });

            if (formulaCells.length > 0) {
                context += `\nFORMULAS IN SHEET:\n`;
                formulaCells.slice(0, 20).forEach((f) => {
                    context += `  ${f}\n`;
                });
                if (formulaCells.length > 20) {
                    context += `  ... and ${formulaCells.length - 20} more formulas\n`;
                }
            }
        }

        return context;
    }

    /**
     * Execute a batch of operations from the LLM
     */
    static async executeOperations(operations: any[]): Promise<void> {
        if (!this.isExcelAvailable()) {
            console.log('Simulating Excel operations:', operations);
            return;
        }

        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            for (const op of operations) {
                try {
                    switch (op.action) {
                        case 'setCellValue':
                            const cellScale = sheet.getRange(op.address);
                            cellScale.values = [[op.value]];
                            break;

                        case 'setFormula':
                            const cellFormula = sheet.getRange(op.address);
                            cellFormula.formulas = [[op.formula]];
                            break;

                        case 'format':
                            const range = sheet.getRange(op.address);
                            if (op.format.bold !== undefined) range.format.font.bold = op.format.bold;
                            if (op.format.fill) range.format.fill.color = op.format.fill;
                            if (op.format.color) range.format.font.color = op.format.color;
                            break;

                        case 'createTable':
                            const table = sheet.tables.add(op.address, true);
                            if (op.name) table.name = op.name;
                            break;

                        case 'createChart':
                            // Basic chart implementation
                            const chartSourceRange = sheet.getRange(op.address);
                            const chart = sheet.charts.add(op.chartType || 'ColumnClustered', chartSourceRange, 'Auto');
                            if (op.title) chart.title.text = op.title;
                            break;

                        case 'createWorksheet':
                            try {
                                const newSheet = context.workbook.worksheets.add(op.name);
                                if (op.activate) newSheet.activate();
                            } catch (e) {
                                console.warn(`Sheet ${op.name} might already exist`);
                            }
                            break;

                        case 'activateWorksheet':
                            const sheetToActivate = context.workbook.worksheets.getItem(op.name);
                            sheetToActivate.activate();
                            break;

                        case 'deleteWorksheet':
                            const sheetToDelete = context.workbook.worksheets.getItem(op.name);
                            sheetToDelete.delete();
                            break;

                        case 'autoFit':
                            // Auto-fit columns in range or used range
                            const autoFitRange = op.address ? sheet.getRange(op.address) : sheet.getUsedRange();
                            autoFitRange.getEntireColumn().format.autofitColumns();
                            break;

                        case 'sortRange':
                            const sortRange = sheet.getRange(op.address);
                            const sortFields = [
                                {
                                    key: op.key || 0, // Column index (0-based) relative to range, or 0 if not specified
                                    ascending: op.ascending !== false // Default true
                                }
                            ];
                            sortRange.sort.apply(sortFields, op.matchCase || false, op.hasHeaders || true, op.orientation || "Rows");
                            break;

                        case 'filterRange':
                            // Apply auto-filter to a range
                            const filterRange = sheet.getRange(op.address);
                            sheet.autoFilter.apply(filterRange);
                            break;

                        case 'createPivotTable':
                            try {
                                const sourceSheet = op.sourceSheet ? context.workbook.worksheets.getItem(op.sourceSheet) : sheet;
                                const sourceRange = sourceSheet.getRange(op.sourceAddress);

                                const destSheet = op.destinationSheet ? context.workbook.worksheets.getItem(op.destinationSheet) : sheet;
                                // Add Pivot Table - Needs Name, Source, and Destination
                                const pivotTableName = op.name || `PivotTable_${Date.now()}`;
                                const pivotTable = destSheet.pivotTables.add(pivotTableName, sourceRange, op.destinationAddress || "A1");

                                if (op.name) pivotTable.name = op.name;

                                // Configure Hierarchies (Rows, Columns, Data)
                                if (op.rows && Array.isArray(op.rows)) {
                                    op.rows.forEach((r: string) => pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(r)));
                                }
                                if (op.columns && Array.isArray(op.columns)) {
                                    op.columns.forEach((c: string) => pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(c)));
                                }
                                if (op.values && Array.isArray(op.values)) {
                                    op.values.forEach((v: any) => {
                                        // v can be string or object { field: string, function: 'Sum' | 'Count' etc }
                                        const fieldName = typeof v === 'string' ? v : v.field;
                                        const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
                                        if (typeof v !== 'string' && v.function) {
                                            dataHierarchy.summarizeBy = v.function; // e.g. Excel.AggregationFunction.sum
                                        }
                                    });
                                }
                            } catch (e) {
                                console.error("Failed to create Pivot Table", e);
                            }
                            break;

                        case 'insertRows':
                            const insertRowRange = sheet.getRange(op.address);
                            insertRowRange.insert(Excel.InsertShiftDirection.down);
                            break;

                        case 'deleteRows':
                            const deleteRowRange = sheet.getRange(op.address);
                            deleteRowRange.delete(Excel.DeleteShiftDirection.up);
                            break;

                        case 'insertColumns':
                            const insertColRange = sheet.getRange(op.address);
                            insertColRange.insert(Excel.InsertShiftDirection.right);
                            break;

                        case 'deleteColumns':
                            const deleteColRange = sheet.getRange(op.address);
                            deleteColRange.delete(Excel.DeleteShiftDirection.left);
                            break;

                        case 'renameWorksheet':
                            const sheetToRename = context.workbook.worksheets.getItem(op.name || op.oldName);
                            sheetToRename.name = op.newName;
                            break;

                        case 'hideWorksheet':
                            const sheetToHide = context.workbook.worksheets.getItem(op.name);
                            sheetToHide.visibility = Excel.SheetVisibility.hidden;
                            break;

                        case 'unhideWorksheet':
                            const sheetToUnhide = context.workbook.worksheets.getItem(op.name);
                            sheetToUnhide.visibility = Excel.SheetVisibility.visible;
                            break;

                    }
                } catch (e) {
                    console.error(`Failed to execute operation ${op.action} on ${op.address}`, e);
                }
            }

            await context.sync();
        });
    }
}

export default ExcelService;
