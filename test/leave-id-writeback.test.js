// 補休編號回寫功能測試
// Red Phase - TDD 失敗測試

function testLeaveIdWritebackFunctions() {
  console.log('🧪 開始補休編號回寫功能測試...');
  
  // 測試1: 配對完成後應回寫用掉補休編號到加班記錄總表
  test('配對完成後應回寫用掉補休編號到加班記錄總表', function() {
    // 先新增加班記錄
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
    
    // 創建補休申請並配對
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
    assert(result.success === true, '補休配對應該成功');
    
    // 檢查加班記錄總表是否有回寫補休編號
    const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const data = overtimeSheet.getDataRange().getValues();
    
    let foundWriteback = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === 'E001' && row[3] === '2025-01-10') {
        const usedLeaveIds = row[11]; // 用掉補休編號欄位
        if (usedLeaveIds && usedLeaveIds.includes('LV-20250115-E001-1')) {
          foundWriteback = true;
          break;
        }
      }
    }
    
    assert(foundWriteback, '應該在加班記錄總表中找到回寫的補休編號');
  });
  
  // 測試2: 應該同步補休編號到員工個人月份分頁
  test('應該同步補休編號到員工個人月份分頁', function() {
    const employee = getActiveEmployees()[0];
    const overtimeId = 'OT-20250110-E001-1';
    const leaveId = 'LV-20250115-E001-1';
    const date = '2025-01-10';
    const sourceMonth = '1月';
    
    // 呼叫回寫到員工個人表格的函數
    const result = writeLeaveIdToEmployeeSheet(employee.id, employee.fileId, date, sourceMonth, leaveId);
    assert(result === true, '回寫到員工個人表格應該成功');
  });
  
  // 測試3: 多筆補休使用同一加班記錄時應用逗號分隔
  test('多筆補休使用同一加班記錄時應用逗號分隔', function() {
    // 先新增加班記錄
    const overtimeRecord = {
      employeeId: 'E002',
      date: '2025-01-10',
      hours: 8,
      type: '上班日加班',
      note: '測試加班',
      sourceMonth: '1月',
      dayOfWeek: '五'
    };
    addOvertimeRecord(overtimeRecord);
    
    // 第一次補休申請
    const leaveRequest1 = {
      leaveId: 'LV-20250115-E002-1',
      employeeId: 'E002',
      employeeName: '測試員工2',
      applicationDate: '2025-01-15',
      leaveDate: '2025-01-16',
      hours: 3,
      note: '第一次補休'
    };
    
    const result1 = addLeaveRequestToMaster(leaveRequest1);
    assert(result1.success === true, '第一次補休配對應該成功');
    
    // 第二次補休申請
    const leaveRequest2 = {
      leaveId: 'LV-20250116-E002-1',
      employeeId: 'E002',
      employeeName: '測試員工2',
      applicationDate: '2025-01-16',
      leaveDate: '2025-01-17',
      hours: 2,
      note: '第二次補休'
    };
    
    const result2 = addLeaveRequestToMaster(leaveRequest2);
    assert(result2.success === true, '第二次補休配對應該成功');
    
    // 檢查加班記錄總表的補休編號是否用逗號分隔
    const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const data = overtimeSheet.getDataRange().getValues();
    
    let foundMultipleIds = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === 'E002' && row[3] === '2025-01-10') {
        const usedLeaveIds = row[11];
        if (usedLeaveIds && usedLeaveIds.includes(',')) {
          foundMultipleIds = true;
          assert(usedLeaveIds.includes('LV-20250115-E002-1'), '應包含第一個補休編號');
          assert(usedLeaveIds.includes('LV-20250116-E002-1'), '應包含第二個補休編號');
          break;
        }
      }
    }
    
    assert(foundMultipleIds, '多筆補休編號應該用逗號分隔');
  });
  
  // 測試4: 配對失敗時不應回寫補休編號
  test('配對失敗時不應回寫補休編號', function() {
    // 創建沒有可用加班時數的員工補休申請
    const leaveRequest = {
      leaveId: 'LV-20250115-E003-1',
      employeeId: 'E003',
      employeeName: '測試員工3',
      applicationDate: '2025-01-15',
      leaveDate: '2025-01-16',
      hours: 8,
      note: '無可用加班時數'
    };
    
    const result = addLeaveRequestToMaster(leaveRequest);
    assert(result.success === false, '配對應該失敗');
    
    // 檢查是否沒有回寫任何補休編號
    const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const data = overtimeSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === 'E003') {
        const usedLeaveIds = row[11];
        assert(!usedLeaveIds || usedLeaveIds === '', '配對失敗時不應有補休編號回寫');
      }
    }
  });
  
  console.log('✅ 補休編號回寫功能測試完成');
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
function runLeaveIdWritebackTests() {
  try {
    testLeaveIdWritebackFunctions();
    console.log('🎉 所有補休編號回寫測試都通過了！');
    return true;
  } catch (error) {
    console.log(`💥 測試失敗: ${error.message}`);
    return false;
  }
}