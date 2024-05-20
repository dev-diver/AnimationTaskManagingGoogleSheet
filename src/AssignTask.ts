function assignAllPartTask() : void {
  deleteNotWorkerSheets()
  makeWorkerSheets()
  cleanWorkerSheets()
  assignWorkersTask()
}

function assignWorkersTask() : void {
  const files = getWorkerSpreadSheets()
  files.forEach(file=>{
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    assignWorkerTask(spreadsheet)
  })
}

function assignWorkerTask(workerSpreadSheet : Spreadsheet) : void {
  const workerName = workerSpreadSheet.getName().split(' ')[0]
  const workerTaskData = getWorkerTaskData(workerName)
  const dataRange = workerSpreadSheet.getRangeByName('작업자데이터시작');
  dataRange.offset(0,0,workerTaskData.length,dataRange.getNumColumns()).setValues(workerTaskData)
}

function cleanWorkerSheets() : void {
  const files = getWorkerSpreadSheets()
  files.forEach(file=>{
    const workerSpreadsheet = SpreadsheetApp.openById(file.getId())
    cleanWorkerSheet(workerSpreadsheet)
  })
}

function cleanWorkerSheet(spreadSheet :Spreadsheet): void {
  const workerSheet = spreadSheet.getSheetByName('작업');
  if (!workerSheet) {
    console.error(`작업 시트를 찾을 수 없습니다: ${spreadSheet.getName()}`);
    return;
  }

  const startRange = spreadSheet.getRangeByName(workerSheet.getName()+'!작업자데이터시작');
  const syncRange = getSyncRange(startRange)
  syncRange.clear()
}

function assignPartTask(sheet : Sheet) : void {
  const templatePartData = getRangeByName('파트데이터시작');
  let startRow = templatePartData.getRow();
  const startColumn = templatePartData.getColumn();
  const lastColumn = templatePartData.getLastColumn();

  let partDataRange = sheet.getRange(startRow, startColumn, 1, lastColumn - startColumn + 1);
  let record = partDataRange.getValues()[0];
  while (record[0]) {
    assignTask(record);
    startRow += 1;
    partDataRange = sheet.getRange(startRow, startColumn, 1, lastColumn - startColumn + 1);
    record = partDataRange.getValues()[0];
  }
}

function assignTask(record : any[], overwrite : boolean = false) : void {
  const worker = record[FieldOffset.WORKER]
  const numFieldName = '작업자'+FieldName.NUMBER+'필드'
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

    const sameRecordRow = findSameRecordRow(record, workerSpreadsheet, workerSheet, numFieldName)
    if(sameRecordRow!=-1){
      console.log("같은 레코드가 있습니다.")
      if(overwrite){
        overwriteRecord(record, workerSpreadsheet, workerSheet, numFieldName, sameRecordRow)
      }
      return;
    }
    
    const startRange = workerSpreadsheet.getRangeByName(numFieldName);
    const dataStartRow = startRange.getRow() + 1;
    const workerStartColumn = startRange.getColumn();

    const workerCutValues = getColumnValues(workerSheet, dataStartRow, workerStartColumn + 1);
    const cutValue = record[FieldOffset.CUT_NUMBER]
    const insertRow = dataStartRow + findInsertPositionIn(workerCutValues, cutValue);

    insertRecord(record, workerSpreadsheet, workerSheet, numFieldName, insertRow);
  }
}