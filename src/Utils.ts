function getSheetByName(name: string) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`${name} 시트를 찾을 수 없습니다.`);
  }
  return sheet;
}

function getRangeByName(name: string) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(name);
  if (!range) {
    throw new Error(`${name} 이름 범위를 찾을 수 없습니다.`);
  }
  return range;
}

function getRowValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number, column: number): string[] {
  const values = [];
  let cell = sheet.getRange(row, column);
  while (cell.getValue()) {
    values.push(cell.getValue());
    cell = cell.offset(0, 1);
  }
  return values;
}

function getColumnValues(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number, column: number): string[] {
  const values = [];
  let cell = sheet.getRange(row, column);
  while (cell.getValue()) {
    values.push(cell.getValue());
    cell = cell.offset(1, 0);
  }
  return values;
}

function getColumnRange(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number, column: number): GoogleAppsScript.Spreadsheet.Range {
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

function getOrCreateFolderByName(folderName: string): GoogleAppsScript.Drive.Folder {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function getOrCreateSpreadsheetByNameInFolder(folderId: string, fileName: string): GoogleAppsScript.Spreadsheet.Spreadsheet {
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

function getProjectName() {
  const settingSheet = getSheetByName('설정');
  const projectNameRange = getRangeByName('프로젝트명');
  return settingSheet.getRange(projectNameRange.getRow(), projectNameRange.getColumn()+1).getValue();
}

function isNewSpreadSheet(spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet): boolean {
  return spreadSheet.getSheets().length === 1 && spreadSheet.getSheets()[0].getName() === '시트1';
}

function copyProjectScript(sourceScriptId: string, targetSpreadsheetId: string) {
  const url = `https://script.googleapis.com/v1/projects/${sourceScriptId}/content`;
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'get',
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  const content = JSON.parse(response.getContentText());

  const targetUrl = `https://script.googleapis.com/v1/projects/${targetSpreadsheetId}/content`;
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

function copyFileToProject(sourceScriptId: string, targetScriptId: string, fileName: string) {

  const fileContent = getFileContent(sourceScriptId, fileName);
  const manifestFile = getManifestFile(sourceScriptId);

  const targetUrl = `https://script.googleapis.com/v1/projects/${targetScriptId}/content`;
  const targetPayload = {
    files: [
      fileContent,
      manifestFile
    ]
  };
  
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

function getFilesInFolder(folderId){
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  return files;
}

function isWorkerSpreadSheet(file){
  return file.getName().endsWith(' 작업');
}

function getWorkerSpreadSheets(){
  const folderId = getOrCreateFolderByName(getProjectName()).getId();
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

function getPartSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = []
  ss.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.endsWith(' 파트')) {
      result.push(sheet)
    }
  })
  return result
}