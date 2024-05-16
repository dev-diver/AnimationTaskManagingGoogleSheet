function getSheetByName(name: string) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`${name} 시트를 찾을 수 없습니다.`);
  }
  return sheet;
}

function getRangeByName(name: string) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(name);
  if (!range) {
    throw new Error(`${name} 이름 범위를 찾을 수 없습니다.`);
  }
  return range;
}

function getRowValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number, column: number): string[] {
  const values = [];
  let cell = sheet.getRange(row, column);
  while (cell.getValue()) {
    values.push(cell.getValue());
    cell = cell.offset(0, 1);
  }
  return values;
}

function getColumnValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number, column: number): string[] {
  const values = [];
  let cell = sheet.getRange(row, column);
  while (cell.getValue()) {
    values.push(cell.getValue());
    cell = cell.offset(1, 0);
  }
  return values;
}

function getColumnRanges(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number, column: number): GoogleAppsScript.Spreadsheet.Range {
  let cell = sheet.getRange(row, column);
  let i = 0;
  while (cell.getValue()) {
    i++;
    cell = cell.offset(1, 0);
  }
  if (i==0) {
    throw new Error('시작이 빈 열입니다.');
  }
  const range = sheet.getRange(row, column, i, 1);
  return range;
}

export { getSheetByName, getRangeByName, getRowValues, getColumnValues, getColumnRanges };