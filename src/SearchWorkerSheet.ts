function isSameRecord(record : any[], workerSpreadsheet : Spreadsheet, workerSheet : Sheet, startRangeName : string,  insertRow : number, ) : boolean {
  const startRange = workerSpreadsheet.getRangeByName(startRangeName);
  const startRowRange = getRowRange(workerSheet, startRange.getRow(), startRange.getColumn());
  const workerStartColumn = startRange.getColumn();
  const workerLastColumn = startRowRange.getLastColumn();
  
  const rangeToCheck = workerSheet.getRange(insertRow, workerStartColumn, 1, workerLastColumn - workerStartColumn + 1);
  const values = rangeToCheck.getValues()[0];
  //앞의 세 값만 비교함
  for (let i = 1; i < 4; i++) {
    if (values[i] !== record[i]) {
      return false;
    }
  }
  return true
}

function findSameRecordRow(record : any[], workerSpreadsheet : Spreadsheet, workerSheet : Sheet, startRangeName: string) : number {
  const startRange = workerSpreadsheet.getRangeByName(startRangeName);
  const dataStartRow = startRange.getRow() + 1;
  const dataStartColumn = startRange.getColumn();
  const workerCutValueColumn = dataStartColumn + FieldOffset.CUT_NUMBER;
  
  const workerCutValues = getColumnValues(workerSheet, dataStartRow, workerCutValueColumn);
  const cutValue = record[FieldOffset.CUT_NUMBER]
  let compareRow = dataStartRow + findInsertPositionIn(workerCutValues, cutValue) - 1;
  while(workerSheet.getRange(compareRow, workerCutValueColumn).getValue()==cutValue){
    if(isSameRecord(record, workerSpreadsheet, workerSheet, startRangeName, compareRow)){
      return compareRow
    }
    compareRow -= 1
  }
  return -1
}

function findInsertPositionIn(cutValues :string[], compareValue: string) : number {

  let left = 0;
  let right = cutValues.length;
  const compareNum = parseInt(compareValue.split('C')[1])
  while (left < right) {
    const mid = Math.floor((left + right) / 2);
    const midValue = cutValues[mid];
    const midNum = parseInt(midValue.split('C')[1])

    if(midNum <= compareNum){
      left = mid + 1;
    }else{
      right = mid;
    }
  }
  return left
}

function overwriteRecord(record: any[], targetSpreadsheet: Spreadsheet, targetSheet: Sheet, startRangeName: string, targetRow: number): void {
  const startRange = targetSpreadsheet.getRangeByName(startRangeName);
  const workerStartColumn = startRange.getColumn();
  const workerDataRange = targetSheet.getRange(targetRow, workerStartColumn, 1, record.length);
  // 값 복사
  workerDataRange.setValues([record]);

}

function insertRecord(record: any[], workerSpreadsheet: Spreadsheet, workerSheet: Sheet, startRangeName: string, insertRow: number): void {

  // partData에서 필요한 값들을 직접 사용하지 않고, record와 관련된 정보를 직접 사용
  const startRange = workerSpreadsheet.getRangeByName(startRangeName);
  const workerStartColumn = startRange.getColumn();

  const lastDataRow = getLastDataRowInRange(startRange);
  const numRows = lastDataRow - insertRow + 1;

  if (numRows > 0) {
    const rangeToMove = workerSheet.getRange(insertRow, workerStartColumn, numRows, record.length);
    rangeToMove.moveTo(workerSheet.getRange(insertRow + 1, workerStartColumn));
  }

  const workerDataRange = workerSheet.getRange(insertRow, workerStartColumn, 1, record.length);
  // 값 복사
  workerDataRange.setValues([record]);
}