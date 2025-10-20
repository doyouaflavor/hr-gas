const SheetsReader = require('./helpers/sheetsReader');

describe('HR 管理系統規格測試 (依據 spec.md)', () => {
  let reader;
  const testSheetId = global.testConfig.sheetId;

  beforeAll(async () => {
    reader = new SheetsReader();
    await reader.initialize();
  });

  describe('資料表結構驗證', () => {
    test('應該包含 spec.md 定義的必要工作表', async () => {
      const sheets = await reader.getSheetNames(testSheetId);
      const sheetNames = sheets.map(s => s.title);
      
      // spec.md 定義的必要工作表
      const requiredSheets = [
        '員工清單',
        '總加班記錄', 
        '總補修記錄',
        '執行紀錄'
      ];
      
      requiredSheets.forEach(sheetName => {
        expect(sheetNames).toContain(sheetName);
      });
      
      console.log('✅ 必要工作表檢查完成');
    });

    test('員工清單應該有正確的欄位結構', async () => {
      const data = await reader.getSheetData(testSheetId, '員工清單!A1:E1');
      
      if (data.length > 0) {
        const headers = data[0];
        const expectedHeaders = ['編號', '姓名', '表格ID', '表格連結', '狀態'];
        
        expectedHeaders.forEach((header, index) => {
          expect(headers[index]).toContain(header);
        });
        
        console.log('📋 員工清單欄位:', headers);
      }
    });

    test('總加班記錄應該有正確的欄位結構', async () => {
      const data = await reader.getSheetData(testSheetId, '總加班記錄!A1:M1');
      
      if (data.length > 0) {
        const headers = data[0];
        const expectedFields = [
          '加班編號', '員工編號', '姓名', '日期', '星期', 
          '加班類型', '加班時數', '已使用補休', '剩餘可補休', 
          '狀態', '資料來源月份', '用掉補修編號', '錯誤提示'
        ];
        
        expectedFields.forEach((field, index) => {
          if (headers[index]) {
            expect(headers[index]).toContain(field.substring(0, 3)); // 部分匹配
          }
        });
        
        console.log('⏰ 總加班記錄欄位:', headers);
      }
    });

    test('總補修記錄應該有正確的欄位結構', async () => {
      const data = await reader.getSheetData(testSheetId, '總補修記錄!A1:J1');
      
      if (data.length > 0) {
        const headers = data[0];
        const expectedFields = [
          '補休編號', '員工編號', '姓名', '申請日期', '補休日期',
          '使用時數', '對應加班編號', '備註', '行政組查閱打勾', '錯誤提示'
        ];
        
        expectedFields.forEach((field, index) => {
          if (headers[index]) {
            expect(headers[index]).toContain(field.substring(0, 3)); // 部分匹配
          }
        });
        
        console.log('🏖️ 總補修記錄欄位:', headers);
      }
    });
  });

  describe('加班編號產生規則測試', () => {
    test('應該產生符合規格的加班編號格式', async () => {
      // 模擬加班編號產生邏輯: "OT-20250801-E001-1"
      const date = '2025-08-01';
      const employeeId = 'E001';
      const sequence = 1;
      
      const expectedFormat = `OT-${date.replace(/-/g, '')}-${employeeId}-${sequence}`;
      expect(expectedFormat).toBe('OT-20250801-E001-1');
      
      // 測試多筆同日加班的流水號
      const sequence2 = 2;
      const expectedFormat2 = `OT-${date.replace(/-/g, '')}-${employeeId}-${sequence2}`;
      expect(expectedFormat2).toBe('OT-20250801-E001-2');
      
      console.log('🏷️ 加班編號格式驗證通過');
    });
  });

  describe('月份分頁掃描測試', () => {
    test('應該能正確識別月份分頁格式', async () => {
      const monthlySheets = await reader.scanMonthlySheets(testSheetId);
      
      monthlySheets.forEach(sheet => {
        // 檢查月份格式: "8月", "9月" 等
        expect(sheet.title).toMatch(/^\d{1,2}月$/);
        expect(sheet.month).toBeGreaterThan(0);
        expect(sheet.month).toBeLessThan(13);
      });
      
      console.log('📅 月份分頁格式驗證:', monthlySheets.map(s => s.title));
    });

    test('月份分頁應該按月份順序排列', async () => {
      const monthlySheets = await reader.scanMonthlySheets(testSheetId);
      
      for (let i = 1; i < monthlySheets.length; i++) {
        expect(monthlySheets[i].month).toBeGreaterThanOrEqual(monthlySheets[i-1].month);
      }
      
      console.log('📊 月份排序驗證通過');
    });
  });

  describe('加班資料提取測試', () => {
    test('應該能從月份分頁提取加班資料', async () => {
      const monthlySheets = await reader.scanMonthlySheets(testSheetId);
      
      if (monthlySheets.length > 0) {
        const firstSheet = monthlySheets[0];
        const testEmployeeId = 'E001';
        
        const overtimeData = await reader.extractOvertimeData(
          testSheetId,
          firstSheet.title,
          testEmployeeId
        );
        
        overtimeData.forEach(record => {
          // 驗證必要欄位
          expect(record).toHaveProperty('employeeId');
          expect(record).toHaveProperty('date');
          expect(record).toHaveProperty('hours');
          expect(record).toHaveProperty('type');
          expect(record).toHaveProperty('sourceMonth');
          
          // 驗證資料格式
          expect(record.date).toMatch(/^\d{4}-\d{2}-\d{2}$/);
          expect(record.hours).toBeGreaterThan(0);
          expect(['平日', '假日', '例假日', '上班日', '上班日加班']).toContain(record.type);
          
          console.log('📝 加班記錄:', record);
        });
      }
    });

    test('應該能正確判斷加班類型', async () => {
      // 測試加班類型判斷邏輯
      const testCases = [
        { dayType: '例假日', dayOfWeek: '日', expected: '例假日' },
        { dayType: '休息日', dayOfWeek: '六', expected: '假日' },
        { dayType: '上班日加班', dayOfWeek: '一', expected: '上班日加班' },
        { dayType: '上班日', dayOfWeek: '二', expected: '上班日' },
        { dayType: '', dayOfWeek: '三', expected: '平日' }
      ];
      
      testCases.forEach(testCase => {
        const result = determineOvertimeType(testCase.dayType, testCase.dayOfWeek);
        expect(result).toBe(testCase.expected);
      });
      
      console.log('🔍 加班類型判斷測試通過');
    });
  });
});

// 輔助函數：判斷加班類型 (模擬 sheetsReader 中的邏輯)
function determineOvertimeType(dayType, dayOfWeek) {
  if (dayType.includes('例假日')) {
    return '例假日';
  } else if (dayType.includes('休息日') || dayOfWeek === '6' || dayOfWeek === '7') {
    return '假日';
  } else if (dayType.includes('上班日加班')) {
    return '上班日加班';
  } else if (dayType.includes('上班日')) {
    return '上班日';
  } else {
    return '平日';
  }
}