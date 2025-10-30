// 例假日加班處理測試
// Red Phase - TDD 失敗測試

function testHolidayOvertimeProcessing() {
  console.log('🧪 開始例假日加班處理測試...');
  
  // 測試資料
  const holidayOvertimeRecord = {
    employeeId: 'E001',
    date: '2025-01-01',
    hours: 8,
    type: '例假日',
    note: '新年例假日加班',
    sourceMonth: '1月',
    dayOfWeek: '日'
  };
  
  const regularOvertimeRecord = {
    employeeId: 'E001', 
    date: '2025-01-02',
    hours: 4,
    type: '上班日加班',
    note: '平日加班',
    sourceMonth: '1月',
    dayOfWeek: '一'
  };
  
  // 測試1: 例假日加班應標記為不可補休狀態
  test('例假日加班應標記為不可補休狀態', function() {
    // 呼叫 addOvertimeRecord 函數
    const result = addOvertimeRecord(holidayOvertimeRecord);
    
    // 檢查加班記錄總表中的狀態
    const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const data = overtimeSheet.getDataRange().getValues();
    
    // 找到最新新增的記錄
    const lastRow = data[data.length - 1];
    const status = lastRow[9]; // 狀態欄位
    
    // 預期狀態應為 "例假日-僅發加班費"
    assert(status === '例假日-僅發加班費', `預期狀態為 "例假日-僅發加班費", 實際為 "${status}"`);
  });
  
  // 測試2: 例假日加班的剩餘可補休應為0
  test('例假日加班的剩餘可補休應為0', function() {
    addOvertimeRecord(holidayOvertimeRecord);
    
    const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const data = overtimeSheet.getDataRange().getValues();
    const lastRow = data[data.length - 1];
    const remainingHours = lastRow[8]; // 剩餘可補休欄位
    
    assert(remainingHours === 0, `預期剩餘可補休為 0, 實際為 ${remainingHours}`);
  });
  
  // 測試3: 一般加班記錄的剩餘可補休應等於加班時數
  test('一般加班記錄的剩餘可補休應等於加班時數', function() {
    addOvertimeRecord(regularOvertimeRecord);
    
    const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const data = overtimeSheet.getDataRange().getValues();
    const lastRow = data[data.length - 1];
    const remainingHours = lastRow[8];
    
    assert(remainingHours === 4, `預期剩餘可補休為 4, 實際為 ${remainingHours}`);
  });
  
  // 測試4: 補休配對應排除例假日加班記錄
  test('補休配對應排除例假日加班記錄', function() {
    // 先新增例假日加班記錄
    addOvertimeRecord(holidayOvertimeRecord);
    
    // 嘗試配對補休
    const result = allocateLeaveToOvertime('E001', 4);
    
    // 應該失敗，因為只有例假日加班記錄（不可補休）
    assert(result.success === false, '預期配對失敗，但實際成功了');
    assert(result.error.includes('補休時數超過可用加班時數'), `預期錯誤訊息包含 "補休時數超過可用加班時數", 實際為 "${result.error}"`);
  });
  
  console.log('✅ 例假日加班處理測試完成');
}

// 測試輔助函數
function test(description, testFunction) {
  try {
    console.log(`  測試: ${description}`);
    testFunction();
    console.log(`  ✅ 通過: ${description}`);
  } catch (error) {
    console.log(`  ❌ 失敗: ${description} - ${error.message}`);
    throw error; // 重新拋出錯誤，確保測試失敗
  }
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

// 執行測試
function runHolidayOvertimeTests() {
  try {
    testHolidayOvertimeProcessing();
    console.log('🎉 所有例假日加班測試都通過了！');
    return true;
  } catch (error) {
    console.log(`💥 測試失敗: ${error.message}`);
    return false;
  }
}