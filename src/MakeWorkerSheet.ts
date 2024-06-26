function makeWorkerSheets(){
  showLoadingScreen_("Loading")
  _makeWorkerSheets(showLoadingScreen_)
}

function _makeWorkerSheets(updateMessage) : void {
  setMainSpreadsheetIdOnMainSheet()
  try{
    updateMessage("템플릿 파일 생성중")
    let templateFile = checkAndCreateWorkerTemplateSheet()
    let names : string[] = getSelectedNames()
    const activeRange = SpreadsheetApp.getActiveSpreadsheet().getActiveRange()
    const startRow = activeRange.getRow()
    const startColumn = activeRange.getColumn()
    names.forEach((name,i) => {
      if (name) {
        const idRange = activeRange.getSheet().getRange(startRow + i, startColumn + 1)
        updateMessage(`${name} 시트 생성중`)
        const spreadSheetName = name.trim() + " 작업";
        if(idRange.getValue()==""){
          const newSpreadSheetId = copyWorkerSheet(templateFile, spreadSheetName)
          idRange.setValue(newSpreadSheetId)
        }
      }
    });
  }catch (e){
    throw Error(e.message)
  }finally{
    hideLoadingScreen_()
  }
}

function checkAndCreateWorkerTemplateSheet() : File {
  const folderId = getShareDriveFolderId()
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName('작업자 템플릿 파일');
  if(!files.hasNext()){
    return makeTemplateSheet()
  }else{
    let template = files.next()
    setWorkerTemplateSheetId(template.getId());
    return template
  }
}

function makeTemplateSheet() : File {
  const folderId = getShareDriveFolderId();
  const templateSheetName = '작업자 템플릿';
  const templateSheet = getMainSheetByName(templateSheetName);
  const newSpreadSheet = getOrCreateSpreadsheetByNameInFolder(folderId, '작업자 템플릿 파일');
  const scriptId = ScriptApp.getScriptId();
  const newSheetName = '작업';
  
  if(isNewSpreadSheet(newSpreadSheet)) {
    templateSheet.copyTo(newSpreadSheet).setName(newSheetName);
    const defaultSheet = newSpreadSheet.getSheets()[0];
    newSpreadSheet.deleteSheet(defaultSheet);

    setMainSpreadsheetId(newSpreadSheet, getActiveSpreadsheetId());

    const newScriptId = createNewScriptProject(newSpreadSheet.getId());
    const fileName = 'WorkerSheetFunc'; // 복사할 파일 이름
    copyLibrarySettingToProject(scriptId, newScriptId, fileName);

    setWorkerTemplateSheetId(newSpreadSheet.getId());
    return DriveApp.getFileById(newSpreadSheet.getId());
  }
}

function copyWorkerSheet(templateFile: File, name : string) : string {
  var newFile = templateFile.makeCopy(name);
  var newSpreadsheetId = newFile.getId();
  return newSpreadsheetId;
  // var newSpreadsheet = SpreadsheetApp.openById(newSpreadsheetId);
  // const newScriptId = createNewScriptProject(newSpreadsheet.getId());
  // const fileName = 'WorkerSheetFunc'; // 복사할 파일 이름
  // copyLibrarySettingToProject(newScriptId, newScriptId, fileName);
}

function deleteNotWorkerSheets() : void { 
  const projectName = getProjectName();
  const driveId = getShareDriveFolderId()
  const folderId = getOrCreateFolderInSharedDrive(driveId,projectName).getId();
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
  const settingsSheet = getMainSheetByName('설정');
  const partRange = getRangeByName('작업자');
  const startRow = partRange.getRow();
  const startColumn = partRange.getColumn();
  const values = getColumnValues(settingsSheet, startRow+1, startColumn);
  return values
}