function AssignAllPartTask() : void {
  makeWorkerSheet()
  const partSheets = getPartSheets()
  partSheets.forEach(sheet=>{
    AssignPartTask(sheet)
  })
}

function AssignPartTask(sheet : Sheet) : void {
  const templatePartData = getRangeByName('파트데이터시작');
  let startRow = templatePartData.getRow();
  const startColumn = templatePartData.getColumn();
  const lastColumn = templatePartData.getLastColumn();

  let partDataRange = sheet.getRange(startRow, startColumn, 1, lastColumn - startColumn + 1);
  let record = partDataRange.getValues()[0];
  while (record[0]) {
    AssignTask(record);
    startRow += 1;
    partDataRange = sheet.getRange(startRow, startColumn, 1, lastColumn - startColumn + 1);
    record = partDataRange.getValues()[0];
  }
}

function AssignTask(record : any[], overwrite : boolean = false) : void {
  const worker = record[Field.WORKER]
  if(!worker){
    return
  }
  const file = getWorkerSpreadSheets().find(spreadSheet => spreadSheet.getName().includes(worker))

  if(file){
    const workerSpreadsheet = SpreadsheetApp.openById(file.getId())
    const workerSheet = workerSpreadsheet.getSheetByName('작업');
    if (!workerSheet) {
      console.error(`작업 시트를 찾을 수 없습니다: ${file.getName()}`);
      return;
    }

    const sameRecordRow = findSameRecordRow(record, workerSpreadsheet, workerSheet)
    if(sameRecordRow!=-1){
      console.log("같은 레코드가 있습니다.")
      if(overwrite){
        overwriteRecord(record, workerSpreadsheet, workerSheet, sameRecordRow)
      }
      return;
    }
    
    const startRange = workerSpreadsheet.getRangeByName('작업자연번필드');
    const dataStartRow = startRange.getRow() + 1;
    const workerStartColumn = startRange.getColumn();

    const workerCutValues = getColumnValues(workerSheet, dataStartRow, workerStartColumn + 1);
    const cutValue = record[Field.CUT_NUMBER]
    const insertPosition = dataStartRow + findInsertPositionIn(workerCutValues, cutValue);

    insertRecord(record, workerSpreadsheet, workerSheet, insertPosition);
  }
}