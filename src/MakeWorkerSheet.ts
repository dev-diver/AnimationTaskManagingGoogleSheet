function makeWorkerSheets() : void {
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
        
        const newScriptId = createNewScriptProject(spreadSheet.getId());
        const fileName = 'WorkerSheetFunc'; // 복사할 파일 이름
        copyLibrarySettingToProject(scriptId, newScriptId, fileName);

        templateSheet.copyTo(spreadSheet).setName(newSheetName);
        // 기본적으로 생성된 빈 시트를 삭제
        const defaultSheet = spreadSheet.getSheets()[0];
        spreadSheet.deleteSheet(defaultSheet);
      }
    }
  });
}

function deleteNotWorkerSheets() : void { 
  const projectName = getProjectName();
  const folderId = getOrCreateFolderByName(projectName).getId();
  const names = makeWorkerList()
  const files = getFilesInFolder(folderId)
  while(files.hasNext()){
    const file = files.next()
    const name = file.getName().split(' ')[0]
    if(!names.includes(name)){
      DriveApp.getFileById(file.getId()).setTrashed(true)
    }
  }
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