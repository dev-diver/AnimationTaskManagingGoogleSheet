function myTest(){
  Logger.log(getSheetData('종합'));
}

function getSheetData(sheetName: string): string[][] {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet with name ${sheetName} not found`);
  }
  const range = sheet.getDataRange();
  return range.getValues();
}

export { getSheetData };
