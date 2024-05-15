import { getSheetName } from '../Code';// Importing the function to test

// Mocking the global objects in Google Apps Script
const mockGetActiveSpreadsheet = jest.fn();
const mockGetActiveSheet = jest.fn();
const mockGetName = jest.fn();

// Mocking the SpreadsheetApp global object
// @ts-ignore
global.SpreadsheetApp = {
  getActiveSpreadsheet: mockGetActiveSpreadsheet,
};

test('getSheetName returns the active sheet name', () => {
  const sheetName = 'Test Sheet';
  mockGetActiveSpreadsheet.mockReturnValue({
    getActiveSheet: mockGetActiveSheet.mockReturnValue({
      getName: mockGetName.mockReturnValue(sheetName),
    }),
  });

  const result = getSheetName();
  expect(result).toBe(sheetName);
});
