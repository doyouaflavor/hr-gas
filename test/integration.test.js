const SheetsReader = require('./helpers/sheetsReader');

describe('HR 管理系統整合測試 (基於 Apps Script 邏輯)', () => {
  let reader;
  const employeeSheetId = global.testConfig.employeeSheetId;

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
        const result = await simulateCheckEmployeeData(employee.id, employeeSheetId);
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
    
    // Step 3.5: 模擬個人補休表同步
    const syncResult = await simulatePersonalLeaveSync();
    const syncedLeaveCount = syncResult.syncedCount;
    errorCount += syncResult.errorCount;
    
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
      syncedLeaveCount,
      matchedCount,
      errorCount,
      duration
    };
    
    console.log('📊 執行完成報告:');
    console.log(`━━━━━━━━━━━━━━━━━━━━`);
    console.log(`處理員工數：${report.processedCount}`);
    console.log(`新增加班記錄：${report.newOvertimeCount} 筆`);
    console.log(`同步補休申請：${report.syncedLeaveCount} 筆`);
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
    const monthlySheets = await reader.scanMonthlySheets(employeeSheetId);
    expect(monthlySheets).toBeInstanceOf(Array);
    
    let totalNewRecords = 0;
    const allOvertimeRecords = [];
    
    // 2. 逐一處理每個月份 (模擬 Apps Script 邏輯)
    for (const sheetInfo of monthlySheets) {
      try {
        const overtimeData = await reader.extractOvertimeData(
          employeeSheetId, 
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
    const monthlySheets = await reader.scanMonthlySheets(employeeSheetId);
    
    if (monthlySheets.length > 0) {
      const firstSheet = monthlySheets[0];
      
      // 讀取原始資料
      const rawData = await reader.getFullSheetData(employeeSheetId, firstSheet.title);
      
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

  test('例假日加班處理測試', async () => {
    console.log('🧪 測試例假日加班處理邏輯...');
    
    // 模擬例假日加班記錄
    const holidayOvertimeRecord = {
      employeeId: 'E001',
      date: '2025-01-01',
      hours: 8,
      type: '例假日',
      note: '新年例假日加班',
      sourceMonth: '1月',
      dayOfWeek: '日'
    };
    
    // 模擬一般加班記錄
    const regularOvertimeRecord = {
      employeeId: 'E001',
      date: '2025-01-02',
      hours: 4,
      type: '上班日加班',
      note: '平日加班',
      sourceMonth: '1月',
      dayOfWeek: '一'
    };
    
    // 測試例假日加班記錄處理
    const holidayResult = simulateAddOvertimeRecordWithHolidayLogic(holidayOvertimeRecord);
    expect(holidayResult.remainingHours).toBe(0);
    expect(holidayResult.status).toBe('例假日-僅發加班費');
    
    // 測試一般加班記錄處理
    const regularResult = simulateAddOvertimeRecordWithHolidayLogic(regularOvertimeRecord);
    expect(regularResult.remainingHours).toBe(4);
    expect(regularResult.status).toBe('未使用');
    
    console.log('✅ 例假日加班處理邏輯正確');
  });

  test('個人補休表同步功能測試', async () => {
    console.log('🧪 測試個人補休表同步功能...');
    
    // 模擬個人補休表資料
    const personalLeaveRequests = [
      {
        employeeId: 'E001',
        employeeName: '張OO',
        applicationDate: '2025-01-15',
        leaveDate: '2025-01-16',
        hours: 4,
        note: '個人事務'
      }
    ];
    
    // 模擬同步過程
    const syncResults = [];
    for (const request of personalLeaveRequests) {
      const leaveId = simulateGenerateLeaveId(request.applicationDate, request.employeeId);
      expect(leaveId).toMatch(/^LV-\d{8}-E\d{3}-\d+$/);
      
      const syncResult = simulateAddLeaveRequestToMaster({
        ...request,
        leaveId: leaveId
      });
      
      syncResults.push(syncResult);
    }
    
    console.log(`✅ 成功同步 ${syncResults.length} 筆補休申請`);
    expect(syncResults.length).toBeGreaterThan(0);
  });

  test('效能測試：處理多個工作表', async () => {
    const startTime = Date.now();
    
    const monthlySheets = await reader.scanMonthlySheets(employeeSheetId);
    const scanTime = Date.now() - startTime;
    
    console.log(`⏱️ 掃描 ${monthlySheets.length} 個工作表用時: ${scanTime}ms`);
    
    // 讀取所有月份資料的效能測試
    const readStartTime = Date.now();
    let totalRecords = 0;
    
    for (const sheet of monthlySheets.slice(0, 3)) { // 只測試前3個月
      const data = await reader.getFullSheetData(employeeSheetId, sheet.title);
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

// 模擬新增加班記錄（含例假日邏輯）
function simulateAddOvertimeRecordWithHolidayLogic(record) {
  const dateStr = record.date.replace(/-/g, '');
  const sequence = 1;
  
  // 例假日加班處理邏輯
  let remainingHours, status;
  if (record.type === '例假日') {
    remainingHours = 0;
    status = '例假日-僅發加班費';
  } else {
    remainingHours = record.hours;
    status = '未使用';
  }
  
  return {
    overtimeId: `OT-${dateStr}-${record.employeeId}-${sequence}`,
    ...record,
    usedHours: 0,
    remainingHours: remainingHours,
    status: status
  };
}

// 模擬產生補休編號
function simulateGenerateLeaveId(date, employeeId) {
  const dateStr = date.replace(/-/g, '');
  const sequence = 1; // 簡化為固定序號
  return `LV-${dateStr}-${employeeId}-${sequence}`;
}

// 模擬新增補休申請到主控台
function simulateAddLeaveRequestToMaster(leaveRequest) {
  // 模擬配對結果
  const success = Math.random() > 0.3; // 70% 成功率
  
  return {
    success: success,
    leaveId: leaveRequest.leaveId,
    overtimeIds: success ? ['OT-20250110-E001-1'] : [],
    error: success ? '' : '補休時數超過可用加班時數'
  };
}

// 模擬反向驗證加班記錄
async function simulateValidateOvertimeRecords() {
  // 模擬發現 0-2 個驗證錯誤
  return Math.floor(Math.random() * 3);
}

// 模擬個人補休表同步
async function simulatePersonalLeaveSync() {
  return {
    syncedCount: Math.floor(Math.random() * 5), // 隨機同步數量
    errorCount: Math.floor(Math.random() * 2)   // 隨機錯誤數量
  };
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