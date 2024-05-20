function doGet(e) {

  let userEmail = getUserEmail();
  
  const worker = e.parameter.worker;
  let allowedEmails;

  if (worker) {

    allowedEmails = getWorkerAuth(worker); 
    if (!allowedEmails.includes(userEmail)) {
      return HtmlService.createHtmlOutput(`<h1>권한이 없습니다.</h1><p>메일:${userEmail}</p><p>관리자한테 문의하세요.</p>`);
    }
    
    const template = HtmlService.createTemplateFromFile('HTML/WorkerData');
    template.worker = worker; // worker 변수를 템플릿에 전달
    return template
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle(worker + ' 작업자 데이터');
  } else {
    allowedEmails = getManagerAuth();
    if (!allowedEmails.includes(userEmail)) {
      return HtmlService.createHtmlOutput(`<h1>권한이 없습니다.</h1><p>메일:${userEmail}</p><p>관리자한테 문의하세요.</p>`);
    }

    const template = HtmlService.createTemplateFromFile('HTML/Index');
    const deploymentUrl = ScriptApp.getService().getUrl();
    template.deploymentUrl = deploymentUrl;
    return template
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('작업자 선택');
  }
}

function getAuthList(pageName: string): string[] {
  const startRange =  getRangeByName('권한').offset(0,1)
  const authValues = getColumnValues(startRange.getSheet(), startRange.getRow(), startRange.getColumn())
  const authIndex = authValues.findIndex(value => value === pageName)
  if(authIndex === -1){
    return []
  }
  const authRow = startRange.getRow() + authIndex
  const mailValues = getRowValues(startRange.getSheet(), authRow, startRange.getColumn() + 1)
  return mailValues
}

function getManagerAuth() : string[] {
  return getAuthList('메인')
}

function getWorkerAuth(worker : string) : string[] {
  return getAuthList(worker)
}

function getWorkerData(worker: string): any[][] {
  // Sample data for illustration; replace with actual data fetching logic
  return [
    ['컷1', '파트1', '2023-01-01', '2023-01-05', '5일', '2023-01-06', '진행중'],
    ['컷2', '파트2', '2023-02-01', '2023-02-05', '5일', '2023-02-06', '완료'],
    ['컷3', '파트3', '2023-03-01', '2023-03-05', '5일', '2023-03-06', '대기중']
  ];
}

function getWorkerData2(worker: string): any[][] {
  const workerSpreadSheets = getWorkerSpreadSheets();
  const file = workerSpreadSheets.find(spreadSheet => spreadSheet.getName().includes(worker));
  if (!file) {
    return [];
  }

  const workerSpreadsheet = SpreadsheetApp.openById(file.getId());
  const workerSheet = workerSpreadsheet.getSheetByName('작업');
  if (!workerSheet) {
    return [];
  }

  const startRange = workerSpreadsheet.getRangeByName(workerSheet.getName()+'!작업자데이터시작');
  const SyncData = getSyncData(startRange);
  return SyncData;
}

function getUserEmail() {
  var url = 'https://people.googleapis.com/v1/people/me?personFields=emailAddresses';
  var params = {
    method: 'GET',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, params);
  var result = JSON.parse(response.getContentText());
  var email = result.emailAddresses ? result.emailAddresses[0].value : 'No email found';
  return email;
}
