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
  const lastColumn = sheet.getLastColumn();

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
    const workerLastRow = workerSheet.getLastRow();
    const workerDataRange = workerSheet.getRange(workerLastRow + 1, 1, 1, lastColumn - startColumn + 1);
    // 값 복사
    const values = partData.getValues();
    workerDataRange.setValues(values);
  }
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