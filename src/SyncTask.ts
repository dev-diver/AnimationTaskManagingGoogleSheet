
function SyncPartDataToWorker() : void {
  const activeSheet = getSpreadsheet();
  const startRange = getRangeByName(activeSheet.getName()+'!파트데이터시작');
  const SyncData = getSyncData(startRange)
  SyncData.forEach(data=>{
    assignTask(data,true)
  })
}

function SyncWorkerToPart() : void {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const activeSheet = ss.getActiveSheet()
  const startRange = ss.getRangeByName(activeSheet.getName()+'!작업자데이터시작');
  const syncData = getSyncData(startRange)

  syncData.forEach(data=>{
    syncAndReportWorkerTask(data)
  })
}

function getSyncRange(startRange: Range) : Range {
  const dataStartRow = startRange.getRow()
  const dataStartColumn = startRange.getColumn()
  const dataEndRow = getLastDataRowInRange(startRange)
  const syncRange = startRange.getSheet().getRange(dataStartRow, dataStartColumn, dataEndRow, startRange.getLastColumn()-dataStartColumn+1)
  return syncRange
}

function getSyncData(startRange : Range) : any[][]{
  const syncRange = getSyncRange(startRange)
  return syncRange.getValues()
}

function syncAndReportWorkerTask(record : any[]){

  const part = record[FieldOffset.PART]
  if(!part){
    console.log("파트가 없습니다.")
    return;
  }

  const manageSpreadsheet = getSpreadsheet()
  const partSheet = getSheetByName(part + " 파트")
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
  const manageSpreadsheet = getSpreadsheet()
  const reportSheet = getSheetByName("로그")
  const numFieldName = '로그'+FieldName.NUMBER+'필드'
  const insertRow = reportSheet.getLastRow() + 1
  insertRecord(record, manageSpreadsheet, reportSheet, numFieldName, insertRow)
}

function getWorkerTaskData(worker: string): any[][] {

  const settingsSheet = getSheetByName('설정');
  const partRange = getRangeByName('파트시작');

  const startRow = partRange.getRow();
  const startColumn = partRange.getColumn();
  const values = getRowValues(settingsSheet, startRow, startColumn + 1);

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
      const sheet = getSheetByName(newSheetName);
      if (sheet) {
        const startRange = getRangeByName(sheet.getName()+'!파트데이터시작');
        const SyncData = getSyncData(startRange)
        SyncData.forEach(data=>{
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