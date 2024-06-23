function assignAllPartTask(){
  showLoadingScreen_("Loading")
  _assignAllPartTask(showLoadingScreen_)
}

function _assignAllPartTask(messageUpdate) : void {
  try{
    // deleteNotWorkerSheets()
    // makeWorkerSheets()
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeRange = ss.getActiveRange()
    if(activeRange.getColumn()!=2 || activeRange.getNumColumns()!=1){
      throw Error("작업자 이름을 선택하고 실행해주세요")
    }
    const names = activeRange.getValues().map((row) : string =>row[0])
    console.log(names)

    if(!getWorkerSpreadSheetId(names[0])){
      setSheetIdToWorkers()
    }
    
    names.forEach(name=>{
      messageUpdate(name+"배치중")
      const spreadSheetId = getWorkerSpreadSheetId(name)
      const spreadSheet = SpreadsheetApp.openById(spreadSheetId)
      // Utilities.sleep(2000);
      cleanWorkerSheet(spreadSheet)
      assignWorkersTask(name, spreadSheet)
    })
  }catch (e){
    messageUpdate(e)
  }finally{
    hideLoadingScreen_()
  }
}


function setSheetIdToWorkers() : void {
  const files = getWorkerSpreadSheets()
  files.forEach(file=>{
    const name = file.getName().split(' ')[0]
    console.log("찾은 이름", name)
    const workerRange = findRangeByWorkerName(name)
    workerRange.offset(0,1).setValue(file.getId())  
  })
}

function findRangeByWorkerName(workerName : string) : Range {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const data = getWorkerNames()
  
  const index = data.indexOf(workerName)
  if(index==-1){
    throw Error("작업자 이름을 찾을 수 없습니다.")
  }

  const startRange = ss.getRangeByName('작업자')
  const row = startRange.getRow() + 1 + index
  const column = startRange.getColumn()
  return startRange.getSheet().getRange(row, column)
}

function getWorkerSpreadSheetId(workerName : string) : string {
  const workerRange = findRangeByWorkerName(workerName)
  return workerRange.offset(0,1).getValue()
}


function applyWorkerSheetFormat(spreadsheet: Spreadsheet){
  fillCheckBox(spreadsheet, '작업자데이터시작', FieldOffset.REPORT);
  fillCheckBox(spreadsheet, '작업자데이터시작', FieldOffset.ALARM);
  // copyColumnFormats(spreadsheet, spreadsheet,'작업자데이터시작', '작업자데이터시작');
  const progressRange = getRangeByName('진행상태');
  const startRow = progressRange.getRow()+1;
  const dataColumn = progressRange.getColumn();
  const dropdownInfoRange = getColumnRange(getMainSheetByName("설정"), startRow, dataColumn);

  const sheet = spreadsheet.getSheetByName('작업');
  const applyFieldRange = getRangeByName("작업자진행현황필드")
  const dataRow = applyFieldRange.getRow() + 1;
  const rowCount = getColumnValues(sheet, dataRow, applyFieldRange.getColumn()).length
  if(rowCount!=0){
    const applyRange = sheet.getRange(dataRow, applyFieldRange.getColumn(), rowCount);
    applyDropdown(dropdownInfoRange, applyRange)
  }
}

function assignWorkersTask(workerName: string, workerSpreadSheet : Spreadsheet) : void {
    assignWorkerTask(workerName, workerSpreadSheet)
    applyWorkerSheetFormat(workerSpreadSheet)
}

function assignWorkerTask(workerName: string, workerSpreadSheet: Spreadsheet) : void {
  const workerTaskData = getWorkerTaskData(workerName)

  const targetRange = workerSpreadSheet.getRangeByName('작업자데이터시작');
  if(workerTaskData.length!=0){
    targetRange.offset(0,0,workerTaskData.length,targetRange.getNumColumns()).setValues(workerTaskData)
  }
}

function cleanWorkerSheets(files : File[]) : void {
  files.forEach(file=>{
    const spreadSheet = SpreadsheetApp.openById(file.getId())
    cleanWorkerSheet(spreadSheet)
  })
}

function cleanWorkerSheet(spreadSheet : Spreadsheet): void {
  const workerSheet = spreadSheet.getSheetByName('작업');
  if (!workerSheet) {
    console.error(`작업 시트를 찾을 수 없습니다: ${spreadSheet.getName()}`);
    return;
  }

  const startRange = spreadSheet.getRangeByName(workerSheet.getName()+'!작업자데이터시작');
  const syncRange = getDataRange(startRange)
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