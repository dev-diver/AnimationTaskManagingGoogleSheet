declare namespace Utils {
    export function getSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet;
    export function getRangeByName(name: string): GoogleAppsScript.Spreadsheet.Range;
    export function getRowValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, startRow: number, startColumn: number): any[];
    export function getColumnValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, startRow: number, startColumn: number): any[];
    export function getColumnRange(sheet: GoogleAppsScript.Spreadsheet.Sheet, startRow: number, startColumn: number): GoogleAppsScript.Spreadsheet.Range;
}