function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  const sheet = e.range.getSheet();
  const editedRange = e.range;
  const editedValue = e.value;
  
  // 예제: 체크박스를 업데이트할 범위 설정
  const sheetName = sheet.getName();
  if(sheetName.endsWith('파트')){
    const startRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(sheetName+'!파트데이터시작');
    let targetRange = getDataRange(startRange)
    targetRange = targetRange.offset(0, 0, targetRange.getNumRows(), targetRange.getNumColumns()-1);
    const checkboxColumnIndex = 13; // 체크박스가 위치한 열의 인덱스 (예: C열이면 3)
    
    // 변경된 셀이 targetRange 내에 있는지 확인
    if (RangeIntersect_(editedRange, targetRange)) {
      // 변경된 셀의 행 번호를 가져옴
      const row = editedRange.getRow();
      
      // 체크박스 열의 셀을 가져와서 TRUE로 설정
      sheet.getRange(row, checkboxColumnIndex).setValue(true);
    }
  }else if(sheetName.endsWith('작업')){
    const startRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(sheetName+'!작업자데이터시작');
    const targetRange = getDataRange(startRange)
    const checkboxColumnIndex = 12; 
    
    // 변경된 셀이 targetRange 내에 있는지 확인
    if (RangeIntersect_(editedRange, targetRange)) {
      // 변경된 셀의 행 번호를 가져옴
      const row = editedRange.getRow();
      // 체크박스 열의 셀을 가져와서 TRUE로 설정
      sheet.getRange(row, checkboxColumnIndex).setValue(true);
    }
  }
}