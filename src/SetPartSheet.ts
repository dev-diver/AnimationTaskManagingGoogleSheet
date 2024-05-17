function applyPart(){
  createSheetsFromSettings();
  performAdditionalTasks();
}

function createSheetsFromSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const settingsSheet = getSheetByName('설정');
  const templateSheet = getSheetByName('파트 템플릿');
  const partRange = getRangeByName('파트시작');

  const startRow = partRange.getRow();
  const startColumn = partRange.getColumn();
  const values = getRowValues(settingsSheet, startRow, startColumn+1);
  const partSheetNames = values.map(part => part.trim() + ' 파트')

  //파트 만들기
  partSheetNames.forEach(sheetName => {
    if (sheetName) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = templateSheet.copyTo(ss).setName(sheetName);
      }
    }
  });

  // 파트에 없는 파트 시트 삭제
  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.endsWith(' 파트') && !partSheetNames.includes(sheetName)) {
      ss.deleteSheet(sheet);
    }
  })
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
        updateWorkerDropdown(startColumn + 1 + i);
        updateProgressDropdown(newSheetName);
      }else{
        throw Error('파트 시트가 존재하지 않습니다.')
      }
    }
  });
}
