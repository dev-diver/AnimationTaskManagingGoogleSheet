function getShareDriveFolderId(){
  let folderId = getRangeByName('폴더ID').offset(0,1).getValue()
  return folderId;
}

function getMainSheetByName(name: string) : Sheet {
  const spreadsheetId = getMainSpreadsheetId();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if(ss.getId()!=spreadsheetId){
    ss = SpreadsheetApp.openById(spreadsheetId);
  }
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`${name} 시트를 찾을 수 없습니다.`);
  }
  return sheet;
}

function getRangeByName(name: string) : Range {
  const ss = getMainSpreadsheet();
  const range = ss.getRangeByName(name);
  if (!range) {
    throw new Error(`${name} 이름 범위를 찾을 수 없습니다.`);
  }
  return range;
}


function getRowValues(sheet: Sheet, row: number, column: number): string[] {
  const values = [];
  let cell = sheet.getRange(row, column);
  while (cell.getValue()) {
    values.push(cell.getValue());
    cell = cell.offset(0, 1);
  }
  return values;
}

function getColumnValues(sheet: Sheet, row: number, column: number): string[] {
  const values = [];
  let cell = sheet.getRange(row, column);
  while (cell.getValue()) {
    values.push(cell.getValue());
    cell = cell.offset(1, 0);
  }
  return values;
}

function getColumnRange(sheet: Sheet, row: number, column: number): Range {
  let cell = sheet.getRange(row, column);
  let i = 0;
  while (cell.getValue()) {
    i++;
    cell = cell.offset(1, 0);
  }
  if (i==0) {
    throw new Error('시작이 빈 열입니다.');
  }
  const range = sheet.getRange(row, column, i, 1);
  return range;
}

function getRowRange(sheet: Sheet, row: number, column: number): Range {
  let cell = sheet.getRange(row, column);
  let i = 0;
  while (cell.getValue()) {
    i++;
    cell = cell.offset(0, 1);
  }
  if (i==0) {
    throw new Error('시작이 빈 행입니다.');
  }
  const range = sheet.getRange(row, column, 1, i);
  return range;
}

function getLastDataRowInRange(range: Range): number {
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const sheetLastRow = sheet.getLastRow();
  const startColumn = range.getColumn();
  const numColumns = 1
  let lastDataRow = startRow;

  for (let row = startRow; row <= sheetLastRow; row++) {
    const rowData = sheet.getRange(row, startColumn, 1, numColumns).getValues()[0];
    const isRowEmpty = rowData.every(cell => cell === '' || cell === null);
    if (!isRowEmpty) {
      lastDataRow = row;
    }
  }

  return lastDataRow;
}

function RangeIntersect_(R1 : Range, R2 : Range) {
  return (R1.getLastRow() >= R2.getRow()) && (R2.getLastRow() >= R1.getRow()) && (R1.getLastColumn() >= R2.getColumn()) && (R2.getLastColumn() >= R1.getColumn());
}

function getCutCount() : number {
  const cutCountRange = getRangeByName("컷수")
  return cutCountRange.offset(0,1).getValue()
}

function getProjectName() : string {
  const settingSheet = getMainSheetByName('설정');
  const projectNameRange = getRangeByName('프로젝트명');
  return settingSheet.getRange(projectNameRange.getRow(), projectNameRange.getColumn()+1).getValue();
}

function getFilesInFolder(folderId : string) : FileIterator{
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  return files;
}

function getOrCreateFolderInSharedDrive(sharedDriveId: string, folderName: string): GoogleAppsScript.Drive.Folder {
  const drive = DriveApp.getFolderById(sharedDriveId);
  const folders = drive.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return drive.createFolder(folderName);
  }
}

function getOrCreateFolderByName(folderName: string): Folder {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function getOrCreateSpreadsheetByNameInFolder(folderId: string, fileName: string): Spreadsheet {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);
  
  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  } else {
    const ns = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(ns.getId());
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file); // 기본 루트 폴더에서 제거
    return ns;
  }
}

function isNewSpreadSheet(spreadSheet: Spreadsheet): boolean {
  return spreadSheet.getSheets().length === 1 && spreadSheet.getSheets()[0].getName() === '시트1';
}

function createNewScriptProject(spreadsheetId: string): string {
  const url = 'https://script.googleapis.com/v1/projects';
  const payload = {
    title: `Script for Spreadsheet ${spreadsheetId}`,
    parentId: spreadsheetId
  };
  
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data.scriptId;
}

function copyProjectScript(sourceScriptId: string, targetScriptId: string) {
  const url = `https://script.googleapis.com/v1/projects/${sourceScriptId}/content`;
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  const content = JSON.parse(response.getContentText());

  const targetUrl = `https://script.googleapis.com/v1/projects/${targetScriptId}/content`;
  const targetOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'put',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    payload: JSON.stringify(content)
  };

  const targetResponse = UrlFetchApp.fetch(targetUrl, targetOptions);
  Logger.log(targetResponse.getContentText());
}

function copyLibrarySettingToProject(sourceScriptId: string, targetScriptId: string, fileName: string) {
  const sourceContent = getProjectContent(sourceScriptId);
  
  const fileContent = sourceContent.files.find((file: any) => file.name === fileName);
  if (!fileContent) {
    throw new Error(`File with name "${fileName}" not found in source project ${sourceScriptId}`);
  }

  // source 프로젝트에서 매니페스트 파일 복사
  let manifestFile = sourceContent.files.find((file: any) => file.name === 'appsscript');
  if (!manifestFile) {
    throw new Error(`Manifest file "appsscript.json" not found in source project ${sourceScriptId}`);
  }

  // manifestFile.source = addLibraryToManifest(JSON.parse(manifestFile.source), sourceScriptId);

  const targetPayload = {
    files: [
      fileContent,
      manifestFile
    ]
  };
  
  const targetUrl = `https://script.googleapis.com/v1/projects/${targetScriptId}/content`;
  const targetOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'put',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    payload: JSON.stringify(targetPayload),
    muteHttpExceptions: true
  };

  const targetResponse = UrlFetchApp.fetch(targetUrl, targetOptions);
  Logger.log(targetResponse.getContentText());
}

function getProjectContent(projectId: string): any {
  const url = `https://script.googleapis.com/v1/projects/${projectId}/content`;
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to fetch project content: ${response.getContentText()}`);
  }

  const content = JSON.parse(response.getContentText());
  return { files: content.files, url };
}

function addLibraryToManifest(manifestContent: any, libraryScriptId: string): string {
  if (!manifestContent.dependencies) {
    manifestContent.dependencies = {};
  }
  if (!manifestContent.dependencies.libraries) {
    manifestContent.dependencies.libraries = [];
  }

  manifestContent.dependencies.libraries.push({
    userSymbol: LIBRARY_NAME,
    libraryId: libraryScriptId,
    version: LIBRARY_VERSION, // 라이브러리 버전 설정
    developmentMode: true // 개발 모드 설정
  });

  return JSON.stringify(manifestContent, null, 2);
}

function getFileContent(projectId: string, fileName: string): any {
  const url = `https://script.googleapis.com/v1/projects/${projectId}/content`;
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const content = JSON.parse(response.getContentText());
  const file = content.files.find((f: any) => f.name === fileName);
  if (file) {
    return file;
  } else {
    throw new Error(`File with name "${fileName}" not found`);
  }
}

function getManifestFile(projectId: string): any {
  const url = `https://script.googleapis.com/v1/projects/${projectId}/content`;
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  const content = JSON.parse(response.getContentText());

  const manifestFile = content.files.find((f: any) => f.name === 'appsscript');
  if (manifestFile) {
    return manifestFile;
  } else {
    throw new Error(`Manifest file "appsscript.json" not found in project ${projectId}`);
  }
}

function getPartSheets() : Sheet[]{
  const ss = getMainSpreadsheet();
  const result = []
  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.endsWith(' 파트')) {
      result.push(sheet)
    }
  })
  return result
}

function getWorkerSpreadSheets() : File[]{
  const driveId = getShareDriveFolderId()
  const folderId = getOrCreateFolderInSharedDrive(driveId,getProjectName()).getId();
  const files = getFilesInFolder(folderId);
  const result = []
  while(files.hasNext()){
    const file = files.next();
    if(isWorkerSpreadSheet(file)){
      result.push(file)
    }
  }
  return result
}

function getWorkerNames() : string[]{
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const startRange = ss.getRangeByName('작업자')
  const data = getColumnValues(startRange.getSheet(), startRange.getRow()+1, startRange.getColumn())
  return data
}

function isWorkerSpreadSheet(file : File) : boolean{
  return file.getName().endsWith(' 작업');
}

function getActiveSpreadsheetId() : string {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function setMainSpreadsheetIdOnMainSheet() {
  const spreadSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  setMainSpreadsheetId(SpreadsheetApp.getActiveSpreadsheet(), spreadSheetId);
}

function setMainSpreadsheetId(spreadsheet : Spreadsheet, mainSpreadsheetId : string) {
  let settingsSheet = spreadsheet.getSheetByName('설정');
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet('설정');
  }
  let cell = spreadsheet.getRangeByName('메인시트ID');
  if(!cell){
    cell = settingsSheet.getRange('A1');
    spreadsheet.setNamedRange('메인시트ID', cell);
  }
  cell.offset(0,1).setValue(mainSpreadsheetId);
}

function getMainSpreadsheetId() : string {
  try{
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadSheetId = activeSheet.getRangeByName('메인시트ID').offset(0,1).getValue();
    if(!spreadSheetId){
      throw new Error('Spreadsheet ID 설정이 안돼 있습니다.');
    } 
    return spreadSheetId;
  }catch(e){
    throw new Error(e);
  }
}

function getMainSpreadsheet() : Spreadsheet{
  return SpreadsheetApp.openById(getMainSpreadsheetId());
}

function getWorkerTemplateSheetId() : string {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  const cell = activeSheet.getRangeByName('템플릿시트ID');
  return cell.offset(0,1).getValue();
}

function setWorkerTemplateSheetId(templateSheetId : string) : void {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  const cell = activeSheet.getRangeByName('템플릿시트ID');
  cell.offset(0,1).setValue(templateSheetId);
}
