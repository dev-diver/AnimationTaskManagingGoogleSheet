//<reference path="type/Utils.d.ts" />
//<reference path="type/SetWorkerDropdown.d.ts" />
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

function performAdditionalTasks() {
  const settingsSheet = getSheetByName('설정');
  const partRange = getRangeByName('파트시작');

  const startRow = partRange.getRow();
  const startColumn = partRange.getColumn();
  const values = getRowValues(settingsSheet, startRow, startColumn + 1);

  values.forEach((part,i) => {
    if (part) {
      const newSheetName = part.trim() + ' 파트';
      const sheet = getSheetByName(newSheetName);
      if (sheet) {
        createWorkerDropdown(startColumn + 1 + i);
        createProgressDropdown(newSheetName);
      }
    }
  });
}
