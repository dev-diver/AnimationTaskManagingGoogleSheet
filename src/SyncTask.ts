
function SyncPartDataToWorker() : void {
  const activeSheet = getSpreadsheet();
  const startRange = getRangeByName(activeSheet.getName()+'!파트데이터시작');
  const SyncData = getSyncData(startRange)
  SyncData.forEach(data=>{
    AssignTask(data,true)
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

function getSyncData(startRange : Range) : any[]{
  const dataStartRow = startRange.getRow()
  const dataStartColumn = startRange.getColumn()
  const syncData = startRange.getSheet().getRange(dataStartRow, dataStartColumn, getCutCount(), startRange.getLastColumn()-dataStartColumn+1)
  return syncData.getValues()
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