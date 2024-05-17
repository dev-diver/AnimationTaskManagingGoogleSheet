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
      const ns = getSpreadsheetByNameInFolder(folderId, spreadSheetName);
      
      if(isNewSpreadSheet(ns)) {
        const newScriptId = createNewScriptProject(ns.getId());
        const fileName = 'Test'; // 복사할 파일 이름
        copyFileToProject(scriptId, newScriptId, fileName);
        const sheet = templateSheet.copyTo(ns).setName(newSheetName);
        // 기본적으로 생성된 빈 시트를 삭제
        const defaultSheet = ns.getSheets()[0];
        ns.deleteSheet(defaultSheet);
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