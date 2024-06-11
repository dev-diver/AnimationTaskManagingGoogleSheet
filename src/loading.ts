
function showLoadingScreen_(msg: string) {
  const template = HtmlService.createTemplateFromFile('HTML/LoadingPage');
  template.msg = msg || '로딩 중...';
  const htmlOutput = template.evaluate()
    .setWidth(400)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Loading');
}

function hideLoadingScreen_() {
  const userInterface = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Loading');
}
