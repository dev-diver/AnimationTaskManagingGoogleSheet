function applyPart(){
  showLoadingScreen_("Loading")
  _applyPart(showLoadingScreen_)
}

function _applyPart(updateMessage) : void {
  try{
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    setMainSpreadsheetId(activeSpreadsheet, activeSpreadsheet.getId());

    updateMessage("파트 생성중")
    createPartSheets();
    updateMessage("파트 꾸미는 중")
    additionalPartSheetTasks();
  }catch(e){
    throw Error(e.message)
  }finally{
    hideLoadingScreen_()
  }
}

function getPartValues(){
  const settingsSheet = getMainSheetByName('설정');
  const partRange = getRangeByName('파트시작');
  const startRow = partRange.getRow();
  const startColumn = partRange.getColumn();
  const values = getColumnValues(settingsSheet, startRow+1, startColumn);
  return values
}

function createPartSheets() : void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const templateSheet = getMainSheetByName('파트 템플릿');
  const values = getPartValues()
  const partSheetNames = values.map(part => part.trim() + ' 파트')

  //파트 만들기
  partSheetNames.forEach(sheetName => {
    if (sheetName) {
      // updateLoadingMessage(sheetName + " 만드는 중")
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = templateSheet.copyTo(ss).setName(sheetName);
      }
    }
  });

  // 파트에 없는 파트 시트 삭제
  // updateLoadingMessage("파트에 없는 시트 삭제중")
  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.endsWith(' 파트') && !partSheetNames.includes(sheetName)) {
      ss.deleteSheet(sheet);
    }
  })
}

function additionalPartSheetTasks() : void {
  const values = getPartValues()
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  values.forEach((part,i) => {
    if (part) {
      // updateLoadingMessage(part + " 꾸미는 중")
      const newSheetName = part.trim() + ' 파트';
      const startRangeName = newSheetName + '!' + FieldName.NUMBER + '필드';
      const sheet = getMainSheetByName(newSheetName);
      if (sheet) {
        sheet.getDataRange().clearDataValidations();
        clearOverCutCount(sheet, startRangeName);
        
        initNumberingData(sheet , startRangeName);
        initPartData(sheet);
        fillTemplateData(sheet, '파트 템플릿!파트데이터시작')
        fillCheckBox(spreadsheet, sheet.getName()+'!파트데이터시작', FieldOffset.REPORT);
        fillCheckBox(spreadsheet, sheet.getName()+'!파트데이터시작', FieldOffset.REPORT+1);
        fillCheckBox(spreadsheet, sheet.getName()+'!파트데이터시작', FieldOffset.REPORT+2);
        copyColumnFormats(spreadsheet, spreadsheet,'파트데이터시작', '파트데이터시작');
        //드롭다운 적용
        updateWorkerDropdown(newSheetName);
        updateProgressDropdown(newSheetName);

        initProgressData(sheet);
      }else{
        throw Error('파트 시트가 존재하지 않습니다.')
      }
    }
  });
}

function fillCheckBox(spreadSheet : Spreadsheet, startRangeName : string, column: number){
  const dataStartRange = spreadSheet.getRangeByName(startRangeName);
  const rowCount = getLastDataRowInRange(dataStartRange) - dataStartRange.getRow() + 1
  const checkBoxRange = dataStartRange.offset(0, column, rowCount, 1)
  
  try {
    // 유효성 검사 규칙 삭제
    checkBoxRange.clearDataValidations();
    
    // 체크박스 삽입
    checkBoxRange.insertCheckboxes();
  } catch (e) {
    console.log("Error occurred: " + e.message);
    
    // 오류가 발생한 경우 해당 셀의 값을 삭제
    checkBoxRange.getValues().forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        const cellRange = checkBoxRange.getCell(rowIndex + 1, colIndex + 1);
        try {
          cellRange.insertCheckboxes();
        } catch (err) {
          cellRange.setValue(''); // 값 삭제
        }
      });
    });
  }
}

function fillTemplateData(sheet: Sheet, startRangeName: string): void {
  const dataRange = getRangeByName(startRangeName);
  const values = dataRange.getValues()[0];
  const formulas = dataRange.getFormulas()[0];
  const startRow = dataRange.getRow();
  const startColumn = dataRange.getColumn();
  const valIndices = [4, 5, 7];
  const funcIndices = [6, 9];

  // getCutCount 함수가 startRow에 기반한 row 수를 반환한다고 가정
  const targetRange = sheet.getRange(startRow, startColumn, getCutCount(), dataRange.getNumColumns());
  const targetValues = targetRange.getValues();

  targetValues.forEach((row,i) => {
    valIndices.forEach(colIndex => {
      if (!row[colIndex]) {
        row[colIndex] = values[colIndex];
      }
    });
    funcIndices.forEach(colIndex => {
      if (!row[colIndex]) {
        let formula = formulas[colIndex].toString(); // 수식을 문자열로 변환

        // 행 번호를 조정하여 수식을 업데이트
        formula = formula.replace(/(\d+)/g, (match) => {
          return (parseInt(match) + i).toString();
        });

        row[colIndex] = `${formula}`; // 수식을 셀에 입력
      }
    });
  });

  targetRange.setValues(targetValues);
}

function copyColumnFormats(sourceSpreadSheet : Spreadsheet, targetSpreadSheet : Spreadsheet, sourceStartRangeName : string, targetStartRangeName : string): void {
  const sourceStartRange = sourceSpreadSheet.getRangeByName(sourceStartRangeName);
  const targetStartRange = targetSpreadSheet.getRangeByName(targetStartRangeName);
  const rowCount = getLastDataRowInRange(targetStartRange) - targetStartRange.getRow() + 1

  for (let col: number = 0; col <= targetStartRange.getNumColumns(); col++) { // B열부터 K열까지 (2열부터 11열까지)
    const sourceRange = sourceStartRange.offset(0, col, 1, 1);
    const targetRange = targetStartRange.offset(0, col, rowCount, 1);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
}

function clearOverCutCount(sheet : Sheet, startRangeName :string) : void {
  const dataRange = getRangeByName('파트데이터시작');
  const startRange = getRangeByName(startRangeName);
  const cutCount = getCutCount();
  const surviveRow = cutCount + startRange.getRow();
  const lastRow = sheet.getLastRow()
  if (lastRow > surviveRow) {
    const clearRange = sheet.getRange(surviveRow+1,dataRange.getColumn(),lastRow-surviveRow,dataRange.getNumColumns()+1)
    clearRange.clear()
    clearRange.clearDataValidations()
  }
}

function initNumberingData(sheet: Sheet, startRangeName: string): void {
  const serialField = getRangeByName(startRangeName);
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

function initPartData(sheet : Sheet) : void {
  const partField = getRangeByName('작업파트필드');
  const partFieldRow = partField.getRow();
  const partFieldColumn = partField.getColumn();
  const partFieldRange = sheet.getRange(partFieldRow+1, partFieldColumn, getCutCount(),1);
  const partName = sheet.getName().replace(' 파트','');
  partFieldRange.setValue(partName);
}

function initProgressData(sheet: Sheet) : void {
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
