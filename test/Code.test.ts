import { getSheetData } from '../src/Code';

// Mocking Google Apps Script services
const mockGetActiveSpreadsheet = jest.fn();
const mockGetSheetByName = jest.fn();
const mockGetDataRange = jest.fn();
const mockGetValues = jest.fn();

global.SpreadsheetApp = {
  getActiveSpreadsheet: mockGetActiveSpreadsheet
} as any;

describe('getSheetData', () => {
  beforeEach(() => {
    jest.clearAllMocks();

    mockGetActiveSpreadsheet.mockReturnValue({
      getSheetByName: mockGetSheetByName
    });
    mockGetSheetByName.mockReturnValue({
      getDataRange: mockGetDataRange
    });
    mockGetDataRange.mockReturnValue({
      getValues: mockGetValues
    });
  });

  test('should return data from the specified sheet', () => {
    const mockData = [
      ['Name', 'Age'],
      ['Alice', 30],
      ['Bob', 25]
    ];
    mockGetValues.mockReturnValue(mockData);

    const data = getSheetData('Sheet1');
    expect(data).toEqual(mockData);

    expect(mockGetActiveSpreadsheet).toHaveBeenCalled();
    expect(mockGetSheetByName).toHaveBeenCalledWith('Sheet1');
    expect(mockGetDataRange).toHaveBeenCalled();
    expect(mockGetValues).toHaveBeenCalled();
  });

  test('should throw an error if the sheet is not found', () => {
    mockGetSheetByName.mockReturnValue(null);

    expect(() => getSheetData('NonExistentSheet')).toThrowError(
      'Sheet with name NonExistentSheet not found'
    );

    expect(mockGetActiveSpreadsheet).toHaveBeenCalled();
    expect(mockGetSheetByName).toHaveBeenCalledWith('NonExistentSheet');
  });
});
