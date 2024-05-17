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
  console.log("StartRow", startRow)
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
  const startColumn = partData.getColumn();
  const lastColumn = partData.getLastColumn();

  if(file){
    console.log(file.getName(), partData.getValues())
    const workerSpreadsheet = SpreadsheetApp.openById(file.getId())
    const workerSheet = workerSpreadsheet.getSheetByName('작업');
    if (!workerSheet) {
      console.error(`작업 시트를 찾을 수 없습니다: ${file.getName()}`);
      return;
    }
    
    const cutValue = partData.getCell(1, 2).getValue();
    const startRange = workerSpreadsheet.getRangeByName('작업자연번필드');
    const workerStartColumn = startRange.getColumn();
    const workerCutValues = getColumnValues(workerSheet, startRange.getRow() + 1, workerStartColumn + 1);
    const insertPosition = findInsertPositionIn(workerCutValues, cutValue) + startRange.getRow() + 1;

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
}

function findInsertPositionIn(cutValues :string[], compareValue: string): number {

  console.log(cutValues)
  let left = 0;
  let right = cutValues.length;
  const compareNum = parseInt(compareValue.split('C')[1])
  while (left < right) {
    const mid = Math.floor((left + right) / 2);
    const midValue = cutValues[mid];
    const midNum = parseInt(midValue.split('C')[1])

    if(midNum < compareNum){
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