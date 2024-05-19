function AssignAllPartTask() {
  makeWorkerSheet()
  const partSheets = getPartSheets()
  partSheets.forEach(sheet=>{
    AssignPartTask(sheet)
  })
}

function AssignPartTask(sheet){
  const templatePartData = getRangeByName('파트데이터시작');
  let startRow = templatePartData.getRow();
  const startColumn = templatePartData.getColumn();
  const lastColumn = templatePartData.getLastColumn();

  let partDataRange = sheet.getRange(startRow, startColumn, 1, lastColumn - startColumn + 1);

  while (partDataRange.getCell(1, 1).getValue()) {
    AssignTask(partDataRange);
    startRow += 1;
    partDataRange = sheet.getRange(startRow, startColumn, 1, lastColumn - startColumn + 1);
  }
}

function AssignTask(partData){
  const worker = partData.getCell(1, 3).getValue();
  if(!worker){
    return
  }
  const file = getWorkerSpreadSheets().find(spreadSheet => spreadSheet.getName().includes(worker))

  if(file){
    console.log(file.getName(), partData.getValues())
    const workerSpreadsheet = SpreadsheetApp.openById(file.getId())
    const workerSheet = workerSpreadsheet.getSheetByName('작업');
    if (!workerSheet) {
      console.error(`작업 시트를 찾을 수 없습니다: ${file.getName()}`);
      return;
    }

    console.log("find same record")
    if(isThereSameRecord(partData, workerSpreadsheet, workerSheet)){
      console.log("같은 레코드가 있습니다.")
      return;
    }
    console.log("no same record")
    
    const startRange = workerSpreadsheet.getRangeByName('작업자연번필드');
    const dataStartRow = startRange.getRow() + 1;
    const workerStartColumn = startRange.getColumn();

    const workerCutValues = getColumnValues(workerSheet, dataStartRow, workerStartColumn + 1);
    const cutValue = partData.getCell(1, 2).getValue();
    const insertPosition = dataStartRow + findInsertPositionIn(workerCutValues, cutValue);

    insertRecord(partData,workerSpreadsheet, workerSheet, insertPosition);
  }
}

function isSameRecord(partData, workerSpreadsheet, workerSheet, insertPosition){

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

function insertRecord(partData, workerSpreadsheet,  workerSheet, insertPosition){

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

function isThereSameRecord(partData, workerSpreadsheet, workerSheet){
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

function findInsertPositionIn(cutValues :string[], compareValue: string): number {

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

function makeWorkerSheet() {
  const templateSheetName = '작업자 템플릿';
  const newSheetName = '작업';
  const templateSheet = getSheetByName(templateSheetName);
  const projectName = getProjectName();
  const folderId = getOrCreateFolderByName(projectName).getId();
  const scriptId = ScriptApp.getScriptId();

  let names : string[] = makeWorkerList()
  names.forEach(name => {
    if (name) {
      const spreadSheetName = name.trim() + " 작업";
      const spreadSheet = getOrCreateSpreadsheetByNameInFolder(folderId, spreadSheetName);
      
      if(isNewSpreadSheet(spreadSheet)) {
        // const newScriptId = createNewScriptProject(spreadSheet.getId());
        // const fileName = 'Test'; // 복사할 파일 이름
        // copyFileToProject(scriptId, newScriptId, fileName);
        templateSheet.copyTo(spreadSheet).setName(newSheetName);
        // 기본적으로 생성된 빈 시트를 삭제
        const defaultSheet = spreadSheet.getSheets()[0];
        spreadSheet.deleteSheet(defaultSheet);
      }
    }
  });
}

function makeWorkerList() : string[]{
  const settingSheet = getSheetByName('설정');
  const partRange = getRangeByName('파트시작');
  const startRow = partRange.getRow();
  const startColumn = partRange.getColumn()+1;

  let field = settingSheet.getRange(startRow, startColumn);
  let names : Set<string> = new Set();
  while(field.getValue()){
    let worker = field.offset(1,0)
    while(worker.getValue()){
      let workerName = worker.getValue()
      names.add(workerName)
      worker = worker.offset(1,0)
    }
    field = field.offset(0,1)
  }
  return [...names]
}