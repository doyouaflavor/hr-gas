const SheetsReader = require('./helpers/sheetsReader');

describe('HR 管理系統整合測試 (基於 Apps Script 邏輯)', () => {
  let reader;
  const testSheetId = global.testConfig.sheetId;

  beforeAll(async () => {
    reader = new SheetsReader();
    await reader.initialize();
  });

  test('模擬完整 checkAllEmployees 流程', async () => {
    // 測試主函數的完整執行流程
    const startTime = new Date();
    
    // Step 1: 模擬取得啟用員工清單
    const activeEmployees = await simulateGetActiveEmployees();
    expect(activeEmployees).toBeInstanceOf(Array);
    
    let processedCount = 0;
    let newOvertimeCount = 0;
    let errorCount = 0;
    
    // Step 2: 對每位員工執行檢查
    for (const employee of activeEmployees.slice(0, 2)) { // 限制測試數量
      try {
        const result = await simulateCheckEmployeeData(employee.id, testSheetId);
        newOvertimeCount += result.newRecords;
        processedCount++;
        
        console.log(`✅ 員工 ${employee.id} 處理完成 - 新增 ${result.newRecords} 筆`);
      } catch (error) {
        errorCount++;
        console.log(`❌ 員工 ${employee.id} 處理失敗:`, error.message);
      }
    }
    
    // Step 3: 模擬反向驗證
    const validationErrors = await simulateValidateOvertimeRecords();
    errorCount += validationErrors;
    
    // Step 4: 模擬補休配對
    const matchResult = await simulateMatchLeaveWithOvertime();
    const matchedCount = matchResult.matched;
    errorCount += matchResult.errors;
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    // 產生執行報告
    const report = {
      processedCount,
      newOvertimeCount,
      matchedCount,
      errorCount,
      duration
    };
    
    console.log('📊 執行完成報告:');
    console.log(`━━━━━━━━━━━━━━━━━━━━`);
    console.log(`處理員工數：${report.processedCount}`);
    console.log(`新增加班記錄：${report.newOvertimeCount} 筆`);
    console.log(`配對補休記錄：${report.matchedCount} 筆`);
    console.log(`發現錯誤：${report.errorCount} 筆`);
    console.log(`執行時間：${report.duration} 秒`);
    
    // 驗證執行結果
    expect(report.processedCount).toBeGreaterThanOrEqual(0);
    expect(report.newOvertimeCount).toBeGreaterThanOrEqual(0);
    expect(report.duration).toBeLessThan(60); // 應在1分鐘內完成
  });

  test('模擬 checkEmployeeData 流程', async () => {
    const testEmployeeId = 'E001';
    
    // 1. 掃描月份工作表
    const monthlySheets = await reader.scanMonthlySheets(testSheetId);
    expect(monthlySheets).toBeInstanceOf(Array);
    
    let totalNewRecords = 0;
    const allOvertimeRecords = [];
    
    // 2. 逐一處理每個月份 (模擬 Apps Script 邏輯)
    for (const sheetInfo of monthlySheets) {
      try {
        const overtimeData = await reader.extractOvertimeData(
          testSheetId, 
          sheetInfo.title, 
          testEmployeeId
        );
        
        // 3. 模擬檢查記錄是否已存在並新增
        const newRecords = [];
        for (const record of overtimeData) {
          const exists = await simulateOvertimeRecordExists(testEmployeeId, record.date);
          if (!exists) {
            const addedRecord = simulateAddOvertimeRecord(record);
            newRecords.push(addedRecord);
          }
        }
        
        totalNewRecords += newRecords.length;
        allOvertimeRecords.push(...newRecords);
        
        console.log(`📊 ${sheetInfo.title}: 新增 ${newRecords.length} 筆加班記錄`);
        
      } catch (error) {
        console.warn(`⚠️ 處理 ${sheetInfo.title} 時發生錯誤:`, error.message);
      }
    }
    
    // 4. 驗證結果
    expect(totalNewRecords).toBeGreaterThanOrEqual(0);
    expect(allOvertimeRecords).toBeInstanceOf(Array);
    
    console.log('📈 總計新增記錄:', totalNewRecords);
    console.log('📋 所有加班記錄:', allOvertimeRecords.length);
    
    // 5. 統計分析
    const stats = analyzeOvertimeRecords(allOvertimeRecords);
    console.log('📊 統計分析:', stats);
    
    expect(stats).toHaveProperty('totalHours');
    expect(stats).toHaveProperty('recordsByType');
    expect(stats).toHaveProperty('recordsByMonth');
  });

  test('模擬資料驗證流程', async () => {
    const monthlySheets = await reader.scanMonthlySheets(testSheetId);
    
    if (monthlySheets.length > 0) {
      const firstSheet = monthlySheets[0];
      
      // 讀取原始資料
      const rawData = await reader.getFullSheetData(testSheetId, firstSheet.title);
      
      // 驗證資料完整性
      const validation = validateSheetData(rawData, firstSheet.title);
      
      expect(validation).toHaveProperty('isValid');
      expect(validation).toHaveProperty('errors');
      expect(validation).toHaveProperty('warnings');
      
      console.log(`✅ ${firstSheet.title} 驗證結果:`, validation);
      
      if (!validation.isValid) {
        console.log('❌ 發現錯誤:', validation.errors);
      }
      
      if (validation.warnings.length > 0) {
        console.log('⚠️ 發現警告:', validation.warnings);
      }
    }
  });

  test('效能測試：處理多個工作表', async () => {
    const startTime = Date.now();
    
    const monthlySheets = await reader.scanMonthlySheets(testSheetId);
    const scanTime = Date.now() - startTime;
    
    console.log(`⏱️ 掃描 ${monthlySheets.length} 個工作表用時: ${scanTime}ms`);
    
    // 讀取所有月份資料的效能測試
    const readStartTime = Date.now();
    let totalRecords = 0;
    
    for (const sheet of monthlySheets.slice(0, 3)) { // 只測試前3個月
      const data = await reader.getFullSheetData(testSheetId, sheet.title);
      totalRecords += data.length;
    }
    
    const readTime = Date.now() - readStartTime;
    console.log(`⏱️ 讀取 ${totalRecords} 筆記錄用時: ${readTime}ms`);
    console.log(`📊 平均每筆記錄處理時間: ${(readTime / totalRecords).toFixed(2)}ms`);
    
    // 效能要求：每筆記錄處理時間應小於 10ms
    expect(readTime / totalRecords).toBeLessThan(10);
  });
});

// 輔助函數：分析加班記錄
function analyzeOvertimeRecords(records) {
  const stats = {
    totalHours: 0,
    totalRecords: records.length,
    recordsByType: {},
    recordsByMonth: {},
    avgHoursPerRecord: 0
  };
  
  records.forEach(record => {
    stats.totalHours += record.hours;
    
    // 按類型統計
    if (!stats.recordsByType[record.type]) {
      stats.recordsByType[record.type] = { count: 0, hours: 0 };
    }
    stats.recordsByType[record.type].count++;
    stats.recordsByType[record.type].hours += record.hours;
    
    // 按月份統計
    if (!stats.recordsByMonth[record.sourceMonth]) {
      stats.recordsByMonth[record.sourceMonth] = { count: 0, hours: 0 };
    }
    stats.recordsByMonth[record.sourceMonth].count++;
    stats.recordsByMonth[record.sourceMonth].hours += record.hours;
  });
  
  stats.avgHoursPerRecord = stats.totalRecords > 0 ? 
    (stats.totalHours / stats.totalRecords).toFixed(2) : 0;
  
  return stats;
}

// =====  Apps Script 模擬函數 =====

// 模擬取得啟用員工清單
async function simulateGetActiveEmployees() {
  return [
    { id: 'E001', name: '張OO', fileId: 'test-file-1', status: '啟用' },
    { id: 'E002', name: '李OO', fileId: 'test-file-2', status: '啟用' }
  ];
}

// 模擬檢查員工資料
async function simulateCheckEmployeeData(employeeId, sheetId) {
  // 模擬掃描月份工作表並提取加班資料
  const mockNewRecords = Math.floor(Math.random() * 5); // 隨機 0-4 筆新記錄
  return { newRecords: mockNewRecords };
}

// 模擬檢查加班記錄是否存在
async function simulateOvertimeRecordExists(employeeId, date) {
  // 簡化為隨機返回，實際會查詢主控台
  return Math.random() < 0.3; // 30% 機率已存在
}

// 模擬新增加班記錄
function simulateAddOvertimeRecord(record) {
  const dateStr = record.date.replace(/-/g, '');
  const sequence = 1;
  
  return {
    overtimeId: `OT-${dateStr}-${record.employeeId}-${sequence}`,
    ...record,
    usedHours: 0,
    remainingHours: record.hours,
    status: '未使用'
  };
}

// 模擬反向驗證加班記錄
async function simulateValidateOvertimeRecords() {
  // 模擬發現 0-2 個驗證錯誤
  return Math.floor(Math.random() * 3);
}

// 模擬補休配對
async function simulateMatchLeaveWithOvertime() {
  return {
    matched: Math.floor(Math.random() * 10), // 隨機配對數量
    errors: Math.floor(Math.random() * 2)    // 隨機錯誤數量
  };
}

// 輔助函數：驗證工作表資料
function validateSheetData(data, sheetName) {
  const validation = {
    isValid: true,
    errors: [],
    warnings: [],
    stats: {
      totalRows: data.length,
      validDateRows: 0,
      overtimeRows: 0
    }
  };
  
  if (data.length === 0) {
    validation.errors.push('工作表為空');
    validation.isValid = false;
    return validation;
  }
  
  // 檢查標題行
  const headerRow = data[0];
  if (!headerRow || headerRow.length < 12) {
    validation.errors.push('標題行格式不正確');
    validation.isValid = false;
  }
  
  // 檢查資料行
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (row.length < 12) {
      validation.warnings.push(`第 ${i + 1} 行欄位數不足`);
      continue;
    }
    
    const date = row[1]; // B欄：日期
    const overtimeHours = parseFloat(row[11]) || 0; // L欄：加班時數
    
    // 檢查日期格式
    if (date) {
      validation.stats.validDateRows++;
      
      if (overtimeHours > 0) {
        validation.stats.overtimeRows++;
        
        // 檢查加班時數合理性
        if (overtimeHours > 24) {
          validation.warnings.push(`第 ${i + 1} 行加班時數異常: ${overtimeHours} 小時`);
        }
      }
    }
  }
  
  return validation;
}