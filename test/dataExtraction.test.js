const SheetsReader = require('./helpers/sheetsReader');

describe('Google Sheets 資料提取測試', () => {
  let reader;
  const testSheetId = global.testConfig.sheetId;

  beforeAll(async () => {
    reader = new SheetsReader();
    await reader.initialize();
  });

  test('應該能掃描月份工作表', async () => {
    const monthlySheets = await reader.scanMonthlySheets(testSheetId);
    
    expect(monthlySheets).toBeInstanceOf(Array);
    
    monthlySheets.forEach(sheet => {
      expect(sheet).toHaveProperty('id');
      expect(sheet).toHaveProperty('title');
      expect(sheet).toHaveProperty('month');
      expect(sheet.title).toMatch(/^\d{1,2}月$/);
      expect(sheet.month).toBeGreaterThan(0);
      expect(sheet.month).toBeLessThan(13);
    });
    
    // 檢查是否按月份排序
    for (let i = 1; i < monthlySheets.length; i++) {
      expect(monthlySheets[i].month).toBeGreaterThanOrEqual(monthlySheets[i-1].month);
    }
    
    console.log('📅 月份工作表:', monthlySheets.map(s => s.title));
  });

  test('應該能讀取指定範圍的資料', async () => {
    // 嘗試讀取第一個月份工作表的前10行
    const monthlySheets = await reader.scanMonthlySheets(testSheetId);
    
    if (monthlySheets.length > 0) {
      const firstSheet = monthlySheets[0];
      const range = `${firstSheet.title}!A1:M10`;
      
      const data = await reader.getSheetData(testSheetId, range);
      
      expect(data).toBeInstanceOf(Array);
      console.log(`📊 ${firstSheet.title} 前10行資料筆數:`, data.length);
      
      if (data.length > 0) {
        console.log('📋 標題行:', data[0]);
      }
    }
  });

  test('應該能提取加班資料', async () => {
    const monthlySheets = await reader.scanMonthlySheets(testSheetId);
    
    if (monthlySheets.length > 0) {
      const firstSheet = monthlySheets[0];
      const testEmployeeId = 'TEST001'; // 使用測試員工ID
      
      const overtimeData = await reader.extractOvertimeData(
        testSheetId, 
        firstSheet.title, 
        testEmployeeId
      );
      
      expect(overtimeData).toBeInstanceOf(Array);
      
      overtimeData.forEach(record => {
        expect(record).toHaveProperty('employeeId');
        expect(record).toHaveProperty('date');
        expect(record).toHaveProperty('hours');
        expect(record).toHaveProperty('type');
        expect(record).toHaveProperty('sourceMonth');
        
        expect(record.employeeId).toBe(testEmployeeId);
        expect(record.hours).toBeGreaterThan(0);
        expect(record.date).toMatch(/^\d{4}-\d{2}-\d{2}$/);
      });
      
      console.log(`⏰ ${firstSheet.title} 加班記錄數:`, overtimeData.length);
      if (overtimeData.length > 0) {
        console.log('📝 第一筆記錄:', overtimeData[0]);
      }
    }
  });

  test('應該能讀取完整工作表資料', async () => {
    const monthlySheets = await reader.scanMonthlySheets(testSheetId);
    
    if (monthlySheets.length > 0) {
      const firstSheet = monthlySheets[0];
      
      const fullData = await reader.getFullSheetData(testSheetId, firstSheet.title);
      
      expect(fullData).toBeInstanceOf(Array);
      console.log(`📊 ${firstSheet.title} 總行數:`, fullData.length);
      
      if (fullData.length > 1) {
        // 檢查第二行是否有日期欄位
        const secondRow = fullData[1];
        if (secondRow.length > 1) {
          console.log('📅 第二行日期欄位:', secondRow[1]);
        }
      }
    }
  });
});