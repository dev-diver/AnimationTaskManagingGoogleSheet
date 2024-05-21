function onEditDo(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  const sheet = e.range.getSheet();
  const editedRange = e.range;
  // const editedValue = e.value;
  
  const sheetName = sheet.getName();
  if(sheetName.endsWith('파트')){
    const startRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(sheetName+'!파트데이터시작');
    let targetRange = getDataRange(startRange)
    const reportColumn = targetRange.getColumn() + FieldOffset.REPORT
    const alarmColumn = targetRange.getColumn() + FieldOffset.ALARM
    targetRange = targetRange.offset(0, FieldOffset.START_DATE, targetRange.getNumRows(), targetRange.getNumColumns()-FieldOffset.START_DATE-1);
    
    // 변경된 셀이 targetRange 내에 있는지 확인
    if (RangeIntersect_(editedRange, targetRange)) {
      // 변경된 셀의 행 번호를 가져옴
      const row = editedRange.getRow();
      
      // 체크박스 열의 셀을 가져와서 TRUE로 설정
      sheet.getRange(row, reportColumn).setValue(true);
      sheet.getRange(row, alarmColumn).setValue(false);
    }
  }else if(sheetName=='작업'){
    // this[LIBRARY_NAME].protectCheck(sheet, editedRange)
    reportCheck(sheet, editedRange)
  }
}