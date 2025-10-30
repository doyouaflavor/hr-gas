// 個人補休表同步功能測試
// Red Phase - TDD 失敗測試

function testPersonalLeaveSyncFunctions() {
  console.log('🧪 開始個人補休表同步功能測試...');
  
  // 測試1: 應該能掃描個人補休表分頁
  test('應該能掃描個人補休表分頁', function() {
    const employees = getActiveEmployees();
    const employee = employees[0]; // 取第一個員工測試
    
    // 呼叫掃描個人補休表函數
    const leaveRequests = scanPersonalLeaveSheet(employee.fileId);
    
    // 應該返回陣列
    assert(Array.isArray(leaveRequests), '掃描結果應該是陣列');
  });
  
  // 測試2: 應該能產生補休編號 LV-YYYYMMDD-E001-1
  test('應該能產生補休編號 LV-YYYYMMDD-E001-1', function() {
    const leaveId = generateLeaveId('2025-01-15', 'E001');
    
    const expectedPattern = /^LV-\d{8}-E\d{3}-\d+$/;
    assert(expectedPattern.test(leaveId), `補休編號格式錯誤: ${leaveId}`);
    assert(leaveId.includes('20250115'), `補休編號應包含日期: ${leaveId}`);
    assert(leaveId.includes('E001'), `補休編號應包含員工編號: ${leaveId}`);
  });
  
  // 測試3: 應該能自動配對可用的加班記錄
  test('應該能自動配對可用的加班記錄', function() {
    // 先新增一筆可用的加班記錄
    const overtimeRecord = {
      employeeId: 'E001',
      date: '2025-01-10',
      hours: 8,
      type: '上班日加班',
      note: '測試加班',
      sourceMonth: '1月',
      dayOfWeek: '五'
    };
    addOvertimeRecord(overtimeRecord);
    
    // 測試補休申請
    const leaveRequest = {
      leaveId: 'LV-20250115-E001-1',
      employeeId: 'E001',
      employeeName: '測試員工',
      applicationDate: '2025-01-15',
      leaveDate: '2025-01-16',
      hours: 4,
      note: '測試補休'
    };
    
    const result = addLeaveRequestToMaster(leaveRequest);
    assert(result.success === true, '補休申請應該成功');
    assert(result.overtimeIds.length > 0, '應該配對到加班記錄');
  });
  
  // 測試4: 應該拒絕例假日加班記錄配對
  test('應該拒絕例假日加班記錄配對', function() {
    // 新增例假日加班記錄
    const holidayRecord = {
      employeeId: 'E002',
      date: '2025-01-01',
      hours: 8,
      type: '例假日',
      note: '新年例假日加班',
      sourceMonth: '1月',
      dayOfWeek: '日'
    };
    addOvertimeRecord(holidayRecord);
    
    // 嘗試補休申請
    const leaveRequest = {
      leaveId: 'LV-20250115-E002-1',
      employeeId: 'E002',
      employeeName: '測試員工2',
      applicationDate: '2025-01-15',
      leaveDate: '2025-01-16',
      hours: 4,
      note: '測試補休'
    };
    
    const result = addLeaveRequestToMaster(leaveRequest);
    assert(result.success === false, '例假日加班不應允許補休配對');
    assert(result.error.includes('補休時數超過可用加班時數'), `錯誤訊息應指出時數不足: ${result.error}`);
  });
  
  // 測試5: 應該處理補休時數超過可用加班時數的錯誤
  test('應該處理補休時數超過可用加班時數的錯誤', function() {
    // 新增小時數加班記錄
    const smallOvertimeRecord = {
      employeeId: 'E003',
      date: '2025-01-10',
      hours: 2,
      type: '上班日加班',
      note: '短時間加班',
      sourceMonth: '1月',
      dayOfWeek: '五'
    };
    addOvertimeRecord(smallOvertimeRecord);
    
    // 申請超過可用時數的補休
    const leaveRequest = {
      leaveId: 'LV-20250115-E003-1',
      employeeId: 'E003',
      employeeName: '測試員工3',
      applicationDate: '2025-01-15',
      leaveDate: '2025-01-16',
      hours: 8, // 超過可用的2小時
      note: '超量補休測試'
    };
    
    const result = addLeaveRequestToMaster(leaveRequest);
    assert(result.success === false, '超量補休應該失敗');
    assert(result.error.includes('補休時數超過可用加班時數'), `錯誤訊息不正確: ${result.error}`);
  });
  
  // 測試6: 應該在中樞表建立補休記錄並加入核選方框
  test('應該在中樞表建立補休記錄並加入核選方框', function() {
    // 先新增可用加班記錄
    const overtimeRecord = {
      employeeId: 'E004',
      date: '2025-01-10',
      hours: 8,
      type: '上班日加班',
      note: '測試加班',
      sourceMonth: '1月',
      dayOfWeek: '五'
    };
    addOvertimeRecord(overtimeRecord);
    
    const leaveRequest = {
      leaveId: 'LV-20250115-E004-1',
      employeeId: 'E004',
      employeeName: '測試員工4',
      applicationDate: '2025-01-15',
      leaveDate: '2025-01-16',
      hours: 4,
      note: '測試補休'
    };
    
    addLeaveRequestToMaster(leaveRequest);
    
    // 檢查補休申請總表
    const leaveSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_DETAILS);
    const data = leaveSheet.getDataRange().getValues();
    const lastRow = data[data.length - 1];
    
    assert(lastRow[0] === 'LV-20250115-E004-1', `補休編號不正確: ${lastRow[0]}`);
    assert(lastRow[1] === 'E004', `員工編號不正確: ${lastRow[1]}`);
    assert(lastRow[8] === false, '行政組查閱打勾應為false'); // 核選方框
  });
  
  // 測試7: 應該回寫補休編號到個人補休表
  test('應該回寫補休編號到個人補休表', function() {
    const employee = getActiveEmployees()[0];
    const leaveId = 'LV-20250115-E001-1';
    const rowIndex = 2; // 假設在第2行
    
    // 呼叫回寫函數
    const result = writeLeaveIdToPersonalSheet(employee.fileId, rowIndex, leaveId);
    assert(result === true, '回寫補休編號應該成功');
  });
  
  console.log('✅ 個人補休表同步功能測試完成');
}

// 測試輔助函數
function test(description, testFunction) {
  try {
    console.log(`  測試: ${description}`);
    testFunction();
    console.log(`  ✅ 通過: ${description}`);
  } catch (error) {
    console.log(`  ❌ 失敗: ${description} - ${error.message}`);
    throw error;
  }
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

// 執行測試
function runPersonalLeaveSyncTests() {
  try {
    testPersonalLeaveSyncFunctions();
    console.log('🎉 所有個人補休表同步測試都通過了！');
    return true;
  } catch (error) {
    console.log(`💥 測試失敗: ${error.message}`);
    return false;
  }
}