function isSameRecord(partData : Range, workerSpreadsheet : Spreadsheet, workerSheet : Sheet, insertPosition : number) : boolean {

  const startRange = workerSpreadsheet.getRangeByName('작업자연번필드');
  const startRowRange = getRowRange(workerSheet, startRange.getRow(), startRange.getColumn());
  const workerStartColumn = startRange.getColumn();
  const workerLastColumn = startRowRange.getLastColumn();
  
  const rangeToCheck = workerSheet.getRange(insertPosition, workerStartColumn, 1, workerLastColumn - workerStartColumn + 1);
  const values = rangeToCheck.getValues()[0];
  //앞의 세 값만 비교함
  for (let i = 1; i < 4; i++) {
    if (values[i] !== partData.getValues()[0][i]) {
      return false;
    }
  }
  return true
}

function isThereSameRecord(partData : Range, workerSpreadsheet : Spreadsheet, workerSheet :Sheet) : boolean {
  const startRange = workerSpreadsheet.getRangeByName('작업자연번필드');
  const dataStartRow = startRange.getRow() + 1;
  const workerCutValueColumn = startRange.getColumn()+1;
  
  const workerCutValues = getColumnValues(workerSheet, dataStartRow, workerCutValueColumn);
  const cutValue = partData.getCell(1, 2).getValue();
  let comparePosition = dataStartRow + findInsertPositionIn(workerCutValues, cutValue) - 1;
  while(workerSheet.getRange(comparePosition, workerCutValueColumn).getValue()==cutValue){
    console.log("comparePosition", comparePosition, workerSheet.getRange(comparePosition, workerCutValueColumn).getValue())
    if(isSameRecord(partData, workerSpreadsheet, workerSheet, comparePosition)){
      return true
    }
    comparePosition -= 1
  }
  return false
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

function insertRecord(partData : Range, workerSpreadsheet : Spreadsheet,  workerSheet : Sheet, insertPosition : number) : void {

  const startColumn = partData.getColumn();
  const lastColumn = partData.getLastColumn();

  const startRange = workerSpreadsheet.getRangeByName('작업자연번필드');
  const workerStartColumn = startRange.getColumn();

  const lastDataRow = getLastDataRowInRange(startRange);
  const numRows = lastDataRow - insertPosition + 1;
  if (numRows > 0) {
    const rangeToMove = workerSheet.getRange(insertPosition, workerStartColumn, numRows, lastColumn - startColumn + 1);
    rangeToMove.moveTo(workerSheet.getRange(insertPosition + 1, workerStartColumn));
  }

  const workerDataRange = workerSheet.getRange(insertPosition, workerStartColumn, 1, lastColumn - startColumn + 1);

  // 값 복사
  const values = partData.getValues();
  workerDataRange.setValues(values);
}