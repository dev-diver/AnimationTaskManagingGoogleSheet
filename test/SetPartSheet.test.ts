//<reference path="../src/SetPartSheet.ts" />
describe('createSheetsFromSettings', () => {
  let mockSpreadsheetApp: any;
  let mockSpreadsheet: any;
  let mockSettingsSheet: any;
  let mockTemplateSheet: any;
  let mockNamedRange: any;
  let createdSheets: any[];

  beforeEach(() => {
    createdSheets = [];

    mockSettingsSheet = {
      getRange: jest.fn().mockImplementation((row, column) => {
        return {
          getValue: jest.fn().mockImplementation(() => {
            if (row === 1 && column === 2) return '원화';
            if (row === 1 && column === 3) return '배경';
            if (row === 1 && column === 4) return '붓질';
            return '';
          }),
          offset: jest.fn().mockImplementation((rowOffset, colOffset) => {
            column += colOffset;
            row += rowOffset;
            return mockSettingsSheet.getRange(row, column);
          })
        };
      }),
    };

    mockTemplateSheet = {
      copyTo: jest.fn().mockReturnValue({
        setName: jest.fn().mockImplementation(name => {
          createdSheets.push(name);
        })
      }),
      getName: jest.fn(),
    };

    mockNamedRange = {
      getRow: jest.fn().mockReturnValue(1),
      getColumn: jest.fn().mockReturnValue(1),
      getSheet: jest.fn().mockReturnValue(mockSettingsSheet),
    };

    mockSpreadsheet = {
      getSheetByName: jest.fn().mockImplementation(name => {
        if (name === '설정') {
          return mockSettingsSheet;
        } else if (name === '파트 템플릿') {
          return mockTemplateSheet;
        } else {
          return null;
        }
      }),
      getRangeByName: jest.fn().mockImplementation(name => {
        if (name === '파트시작') {
          return mockNamedRange;
        } else {
          return null;
        }
      }),
      insertSheet: jest.fn(),
    };

    mockSpreadsheetApp = {
      getActiveSpreadsheet: jest.fn().mockReturnValue(mockSpreadsheet),
    };

    global.SpreadsheetApp = mockSpreadsheetApp;
  });

  it('should create sheets based on settings by copying template', () => {
    SetPartSheet.createSheetsFromSettings();

    expect(createdSheets).toEqual(['원화 파트', '배경 파트', '붓질 파트']);
  });

  it('should not create duplicate sheets', () => {
    mockSpreadsheet.getSheetByName = jest.fn().mockImplementation(name => {
      if (name === '설정' || name === '원화 파트' || name === '배경 파트' || name === '붓질 파트') {
        return mockSettingsSheet;
      } else if (name === '파트 템플릿') {
        return mockTemplateSheet;
      } else {
        return null;
      }
    });

    SetPartSheet.createSheetsFromSettings();

    expect(createdSheets).toEqual([]);
  });
});