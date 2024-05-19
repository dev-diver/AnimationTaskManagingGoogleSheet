
function SyncPartDataToWorker() : void {
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const startRange = getRangeByName(activeSheet.getName()+'!파트데이터시작');
  const SyncData = getSyncData(startRange)
  SyncData.forEach(data=>{
    AssignTask(data,true)
  })
}

function SyncWorkerToPart() : void {
  const activeSheet = SpreadsheetApp.getActiveSheet()
  const startRange = getRangeByName(activeSheet.getName()+'!작업자데이터시작');
  console.log(getSyncData(startRange))
}

function getSyncData(startRange : Range) : any[]{
  const dataStartRow = startRange.getRow()
  const dataStartColumn = startRange.getColumn()
  const syncData = startRange.getSheet().getRange(dataStartRow, dataStartColumn, getCutCount(), startRange.getLastColumn()-dataStartColumn+1)
  return syncData.getValues()
}