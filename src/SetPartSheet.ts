function applyPart() : void {
  setActiveSpreadsheetId();
  createSheetsFromSettings();
  performAdditionalTasks();
}

function createSheetsFromSettings() : void {
  const ss = getSpreadsheet();

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

function performAdditionalTasks() : void {
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
        //드롭다운 적용
        updateWorkerDropdown(startColumn + 1 + i);
        updateProgressDropdown(newSheetName);
        initNumberingData(sheet);
        initPartData(sheet);
        initProgressData(sheet);
      }else{
        throw Error('파트 시트가 존재하지 않습니다.')
      }
    }
  });
}

function initNumberingData(sheet: Sheet): void {
  const serialField = getRangeByName('연번필드');
  const serialFieldRow = serialField.getRow()+1
  const serialFieldColumn = serialField.getColumn();

  const count = getCutCount(); // 1부터 10까지의 값을 넣을 개수

  // 연번과 코드 값을 채울 범위
  const serialRange = sheet.getRange(serialFieldRow, serialFieldColumn, count, 1);
  const codeRange = sheet.getRange(serialFieldRow, serialFieldColumn + 1, count, 1);

  // 값 배열 생성
  const serialValues = [];
  const codeValues = [];
  for (let i = 1; i <= count; i++) {
    serialValues.push([i]);
    codeValues.push([`C${String(i).padStart(3, '0')}`]); // C001 형식으로 만들기 위해 padStart 사용
  }

  // 값 설정
  serialRange.setValues(serialValues);
  codeRange.setValues(codeValues);
}

function initPartData(sheet : GoogleAppsScript.Spreadsheet.Sheet) : void {
  const partField = getRangeByName('작업파트필드');
  const partFieldRow = partField.getRow();
  const partFieldColumn = partField.getColumn();
  const partFieldRange = sheet.getRange(partFieldRow+1, partFieldColumn, getCutCount(),1);
  const partName = sheet.getName().replace(' 파트','');
  partFieldRange.setValue(partName);
}

function initProgressData(sheet: GoogleAppsScript.Spreadsheet.Sheet) : void {
  const progressField = getRangeByName('진행현황필드');
  const progressFieldRow = progressField.getRow();
  const progressFieldColumn = progressField.getColumn();
  const progressFieldRange = sheet.getRange(progressFieldRow+1, progressFieldColumn, getCutCount(),1);
  
  const values = progressFieldRange.getValues();
  values.forEach((value, index) => {
    if (!value[0]) {
      progressFieldRange.getCell(index+1, 1).setValue('시작전');
    }
  });
}
