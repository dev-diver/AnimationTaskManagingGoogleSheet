
function syncPartDataToWorker(){
  showLoadingScreen_("Loading")
  _syncPartDataToWorker(showLoadingScreen_)
}

function _syncPartDataToWorker(updateMessage) : void {
  try{
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    const startRange = getRangeByName(activeSheet.getName()+'!파트데이터시작');
    const syncData = getSyncData(startRange, (row : any[])=>{
      if(row[FieldOffset.REPORT] === true){
        row[FieldOffset.REPORT] = false
      }
      return row
    })
    syncData.forEach(data => {
      assignTask(data,true)
    })
  }finally{

    hideLoadingScreen_()
  }
}

function syncWorkerToPart() : void {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const activeSheet = ss.getActiveSheet()
  const startRange = ss.getRangeByName(activeSheet.getName()+'!작업자데이터시작');
  const date = new Date()//Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy. M. dd');

  const syncData = getSyncData(startRange, (row : any[]) => {
    if(row[FieldOffset.REPORT] === true){
      row[FieldOffset.REPORT_COUNT] += 1
      row[FieldOffset.REPORT_DATE] = date
      row[FieldOffset.REPORT] = false
    }
    return row
  })

  syncData.forEach(data=>{
    syncAndReportWorkerTask(data)
  })
}


function getDataRange(startRange: Range) : Range {
  const dataStartRow = startRange.getRow()
  const dataEndRow = getLastDataRowInRange(startRange)
  const dataRange = startRange.offset(0, 0, dataEndRow - dataStartRow + 1, startRange.getNumColumns())
  return dataRange
}

function getSyncData(startRange : Range, callback: (row : string[]) => any[]) : string[][]{
  const syncRange = getDataRange(startRange)
  const data = syncRange.getValues()
  const newData = data.map(row=>callback([...row]))
  const filteredData = data
    .map((row, i) =>{
      if(row[FieldOffset.REPORT] === true){
        const result = [...newData[i]]
        result[FieldOffset.ALARM] = true
        return result
      }
      return null
    })
    .filter(row => row !== null);
  syncRange.setValues(newData)
  return filteredData
}

function syncAndReportWorkerTask(record : any[]){

  const part = record[FieldOffset.PART]
  if(!part){
    console.log("파트가 없습니다.")
    return;
  }

  const manageSpreadsheet = getMainSpreadsheet()
  const partSheet = getMainSheetByName(part + " 파트")
  if(!partSheet){
    throw new Error("파트 시트가 없습니다.")
  }

  const numFieldName = FieldName.NUMBER+'필드'
  const sameRecordRow = findSameRecordRow(record, manageSpreadsheet, partSheet, numFieldName)
  if(sameRecordRow!=-1){
    overwriteRecord(record, manageSpreadsheet, partSheet, numFieldName, sameRecordRow)
    reportWorkerTask(record)
    return;
  }
}

function reportWorkerTask(record : any[]){
  const manageSpreadsheet = getMainSpreadsheet()
  const reportSheet = getMainSheetByName("로그")
  const numFieldName = '로그'+FieldName.NUMBER+'필드'
  const insertRow = reportSheet.getLastRow() + 1
  insertRecord(record, manageSpreadsheet, reportSheet, numFieldName, insertRow)
}

function getWorkerTaskData(worker: string): any[][] {

  const values = getPartValues()

  let records = []
  const indicies = [
    FieldOffset.CUT_NUMBER,
    FieldOffset.PART,
    FieldOffset.START_DATE,
    FieldOffset.END_DATE,
    FieldOffset.TERM,
    FieldOffset.REPORT_DATE,
    FieldOffset.PROGRESS_STATE,]
  values.forEach((part,i) => {
    if (part) {
      const newSheetName = part.trim() + ' 파트';
      const sheet = getMainSheetByName(newSheetName);
      if (sheet) {
        const startRange = getRangeByName(sheet.getName()+'!파트데이터시작');
        const syncData = getDataRange(startRange).getValues()
        syncData.forEach(data=>{
          if(data[FieldOffset.WORKER] === worker){
            // data = data.map((value) => {
            //   return indicies.map(index => value[index])
            // })
            records.push(data)
          }
        })
      }else{
        throw Error('파트 시트가 존재하지 않습니다.')
      }
    }
  });

  records.sort((a, b) => {
    let aOrd = a[FieldOffset.NUMBER]//a[FieldOffset.CUT_NUMBER].split('C')[1]
    let bOrd = b[FieldOffset.NUMBER]//b[FieldOffset.CUT_NUMBER].split('C')[1]
    return aOrd - bOrd
  })
  return records;
}

function reportCheck(sheet : Sheet, editedRange : Range){
  const sheetName = sheet.getName();
  const startRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(sheetName+'!작업자데이터시작');
  const dataRange = getDataRange(startRange)
  const targetRange = dataRange.offset(0, FieldOffset.PROGRESS_STATE, dataRange.getNumRows(), 1);
  const reportColumn = dataRange.getColumn() + FieldOffset.REPORT
  const alarmColumn = dataRange.getColumn() + FieldOffset.ALARM
    
  // 변경된 셀이 targetRange 내에 있는지 확인
  if (RangeIntersect_(editedRange, targetRange)) {
    // 변경된 셀의 행 번호를 가져옴
    const row = editedRange.getRow();
    // 체크박스 열의 셀을 가져와서 TRUE로 설정
    sheet.getRange(row, reportColumn).setValue(true);
    sheet.getRange(row, alarmColumn).setValue(false);
  }
}