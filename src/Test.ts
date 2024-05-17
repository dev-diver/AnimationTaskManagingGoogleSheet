function Test() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(10, 10).setValue('Hello, World!');
}