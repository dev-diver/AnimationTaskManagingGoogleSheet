function getOrCreateFolderByName(folderName: string): GoogleAppsScript.Drive.Folder {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function getSpreadsheetByNameInFolder(folderId: string, fileName: string): GoogleAppsScript.Spreadsheet.Spreadsheet {
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
      
      if(ns.getSheets().length === 1 && ns.getSheets()[0].getName() === '시트1') {
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