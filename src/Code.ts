function createSheetsFromSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const settingsSheet = getSheetByName('설정');
  const templateSheet = getSheetByName('파트 템플릿');
  const partRange = getRangeByName('파트시작');

  const startRow = partRange.getRow();
  const startColumn = partRange.getColumn();
  const values = getRowValues(settingsSheet, startRow, startColumn+1);
  
  values.forEach(part => {
    if (part) {
      const newSheetName = part.trim() + ' 파트';
      let sheet = ss.getSheetByName(newSheetName);
      if (!sheet) {
        sheet = templateSheet.copyTo(ss).setName(newSheetName);
      }
    }
  });
}

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

export { createSheetsFromSettings, getSheetByName, getRangeByName, getRowValues, getColumnValues };