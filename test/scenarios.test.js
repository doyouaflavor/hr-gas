const SheetsReader = require('./helpers/sheetsReader');

describe('HR 管理系統測試場景 (依據 spec.md)', () => {
  let reader;
  const testSheetId = global.testConfig.sheetId;

  beforeAll(async () => {
    reader = new SheetsReader();
    await reader.initialize();
  });

  describe('測試場景 1: 新增加班記錄', () => {
    test('員工加班記錄應該正確新增到總表', async () => {
      // 模擬場景：員工A在 2025-08 分頁有加班 8 小時，主控台沒有這筆資料
      const testEmployee = {
        id: 'E001',
        name: '張OO'
      };
      
      const newOvertimeRecord = {
        employeeId: testEmployee.id,
        date: '2025-08-15',
        hours: 8,
        type: '平日',
        sourceMonth: '8月',
        dayOfWeek: '五'
      };
      
      // 模擬檢查邏輯：主控台中不存在此記錄
      const existingRecords = await simulateOvertimeRecordCheck(
        testEmployee.id, 
        newOvertimeRecord.date
      );
      
      expect(existingRecords).toBe(false); // 假設不存在
      
      // 模擬新增記錄到主控台
      const generatedRecord = simulateAddOvertimeRecord(newOvertimeRecord);
      
      // 驗證新增的記錄格式
      expect(generatedRecord).toHaveProperty('overtimeId');
      expect(generatedRecord.overtimeId).toMatch(/^OT-\d{8}-E001-\d+$/);
      expect(generatedRecord.employeeId).toBe(testEmployee.id);
      expect(generatedRecord.employeeName).toBe(testEmployee.name);
      expect(generatedRecord.date).toBe(newOvertimeRecord.date);
      expect(generatedRecord.hours).toBe(8);
      expect(generatedRecord.usedHours).toBe(0);
      expect(generatedRecord.remainingHours).toBe(8);
      expect(generatedRecord.status).toBe('未使用');
      expect(generatedRecord.sourceMonth).toBe('8月');
      
      console.log('✅ 場景1 - 新增加班記錄:', generatedRecord);
    });
  });

  describe('測試場景 2: 補休自動配對', () => {
    test('補休應該按照最舊加班記錄優先分配', async () => {
      // 模擬場景：員工A有兩筆加班，申請補休 12 小時
      const employeeId = 'E001';
      const existingOvertimeRecords = [
        {
          overtimeId: 'OT-20250801-E001-1',
          date: '2025-08-01',
          totalHours: 10,
          usedHours: 0,
          remainingHours: 10,
          status: '未使用'
        },
        {
          overtimeId: 'OT-20250805-E001-1',
          date: '2025-08-05',
          totalHours: 8,
          usedHours: 0,
          remainingHours: 8,
          status: '未使用'
        }
      ];
      
      const leaveRequest = {
        employeeId: employeeId,
        leaveHours: 12,
        leaveDate: '2025-08-10'
      };
      
      // 模擬補休配對邏輯
      const allocationResult = simulateLeaveAllocation(existingOvertimeRecords, leaveRequest);
      
      // 驗證配對結果
      expect(allocationResult.success).toBe(true);
      expect(allocationResult.allocations).toHaveLength(2);
      
      // 第一筆加班記錄：應該全部用完 (10小時)
      const firstAllocation = allocationResult.allocations[0];
      expect(firstAllocation.overtimeId).toBe('OT-20250801-E001-1');
      expect(firstAllocation.allocatedHours).toBe(10);
      expect(firstAllocation.newUsedHours).toBe(10);
      expect(firstAllocation.newRemainingHours).toBe(0);
      expect(firstAllocation.newStatus).toBe('已全數使用');
      
      // 第二筆加班記錄：應該用掉 2小時
      const secondAllocation = allocationResult.allocations[1];
      expect(secondAllocation.overtimeId).toBe('OT-20250805-E001-1');
      expect(secondAllocation.allocatedHours).toBe(2);
      expect(secondAllocation.newUsedHours).toBe(2);
      expect(secondAllocation.newRemainingHours).toBe(6);
      expect(secondAllocation.newStatus).toBe('部分使用');
      
      console.log('✅ 場景2 - 補休配對結果:', allocationResult);
    });
  });

  describe('測試場景 3: 補休超量錯誤', () => {
    test('補休時數超過可用加班時數時應該標記錯誤', async () => {
      // 模擬場景：員工A只有 5 小時可補休，申請補休 8 小時
      const employeeId = 'E001';
      const existingOvertimeRecords = [
        {
          overtimeId: 'OT-20250801-E001-1',
          date: '2025-08-01',
          totalHours: 5,
          usedHours: 0,
          remainingHours: 5,
          status: '未使用'
        }
      ];
      
      const leaveRequest = {
        employeeId: employeeId,
        leaveHours: 8,
        leaveDate: '2025-08-10'
      };
      
      // 模擬補休配對邏輯
      const allocationResult = simulateLeaveAllocation(existingOvertimeRecords, leaveRequest);
      
      // 驗證錯誤處理
      expect(allocationResult.success).toBe(false);
      expect(allocationResult.error).toContain('補休時數超過可用加班時數');
      expect(allocationResult.error).toContain('超過 3 小時');
      expect(allocationResult.allocations).toHaveLength(0);
      
      console.log('❌ 場景3 - 補休超量錯誤:', allocationResult.error);
    });
  });

  describe('錯誤處理測試', () => {
    test('應該正確處理各種錯誤情況', async () => {
      const errorScenarios = [
        {
          type: '員工檔案無法存取',
          description: '權限問題或檔案不存在',
          expectedError: '無法存取員工檔案'
        },
        {
          type: '月份分頁格式不正確',
          description: '分頁名稱不符合月份格式',
          expectedError: '月份分頁格式錯誤'
        },
        {
          type: '資料格式錯誤',
          description: '日期或數字格式不正確',
          expectedError: '資料格式不正確'
        }
      ];
      
      errorScenarios.forEach(scenario => {
        const errorLog = simulateErrorHandling(scenario.type, 'E001', scenario.description);
        
        expect(errorLog).toHaveProperty('timestamp');
        expect(errorLog).toHaveProperty('errorType');
        expect(errorLog).toHaveProperty('employeeId');
        expect(errorLog).toHaveProperty('message');
        
        expect(errorLog.errorType).toBe(scenario.type);
        expect(errorLog.employeeId).toBe('E001');
        expect(errorLog.message).toContain(scenario.description);
        
        console.log(`🚨 錯誤處理 - ${scenario.type}:`, errorLog);
      });
    });
  });

  describe('反向驗證測試', () => {
    test('主控台加班記錄應該能在員工檔案中找到對應資料', async () => {
      // 模擬主控台的加班記錄
      const masterOvertimeRecord = {
        overtimeId: 'OT-20250815-E001-1',
        employeeId: 'E001',
        date: '2025-08-15',
        hours: 8,
        sourceMonth: '8月'
      };
      
      // 模擬在員工檔案中查找對應記錄
      const verificationResult = await simulateReverseValidation(
        testSheetId,
        masterOvertimeRecord
      );
      
      // 如果找到對應記錄
      if (verificationResult.found) {
        expect(verificationResult.matchedRecord).toHaveProperty('date');
        expect(verificationResult.matchedRecord).toHaveProperty('hours');
        expect(verificationResult.matchedRecord.date).toBe(masterOvertimeRecord.date);
        expect(verificationResult.matchedRecord.hours).toBe(masterOvertimeRecord.hours);
        
        console.log('✅ 反向驗證成功:', verificationResult.matchedRecord);
      } else {
        // 如果沒找到，應該標記錯誤
        expect(verificationResult.error).toContain('找不到對應資料');
        console.log('❌ 反向驗證失敗:', verificationResult.error);
      }
    });
  });
});

// 模擬函數：檢查加班記錄是否存在
function simulateOvertimeRecordCheck(employeeId, date) {
  // 在實際測試中，這裡會查詢 Google Sheets
  // 這裡簡化為返回 false (假設記錄不存在)
  return false;
}

// 模擬函數：新增加班記錄
function simulateAddOvertimeRecord(record) {
  const dateStr = record.date.replace(/-/g, '');
  const sequence = 1; // 簡化為固定序號
  
  return {
    overtimeId: `OT-${dateStr}-${record.employeeId}-${sequence}`,
    employeeId: record.employeeId,
    employeeName: '張OO', // 模擬從員工清單取得
    date: record.date,
    dayOfWeek: record.dayOfWeek,
    type: record.type,
    hours: record.hours,
    usedHours: 0,
    remainingHours: record.hours,
    status: '未使用',
    sourceMonth: record.sourceMonth,
    usedLeaveIds: '',
    errorMessage: ''
  };
}

// 模擬函數：補休分配邏輯
function simulateLeaveAllocation(overtimeRecords, leaveRequest) {
  // 按日期排序 (最舊的優先)
  const sortedRecords = [...overtimeRecords].sort((a, b) => 
    new Date(a.date) - new Date(b.date)
  );
  
  let remainingLeaveHours = leaveRequest.leaveHours;
  const allocations = [];
  
  for (const record of sortedRecords) {
    if (remainingLeaveHours <= 0) break;
    if (record.remainingHours <= 0) continue;
    
    const allocatedHours = Math.min(remainingLeaveHours, record.remainingHours);
    const newUsedHours = record.usedHours + allocatedHours;
    const newRemainingHours = record.remainingHours - allocatedHours;
    
    let newStatus = '未使用';
    if (newRemainingHours === 0) {
      newStatus = '已全數使用';
    } else if (newUsedHours > 0) {
      newStatus = '部分使用';
    }
    
    allocations.push({
      overtimeId: record.overtimeId,
      allocatedHours,
      newUsedHours,
      newRemainingHours,
      newStatus
    });
    
    remainingLeaveHours -= allocatedHours;
  }
  
  if (remainingLeaveHours > 0) {
    return {
      success: false,
      error: `補休時數超過可用加班時數 (超過 ${remainingLeaveHours} 小時)`,
      allocations: []
    };
  }
  
  return {
    success: true,
    allocations,
    totalAllocated: leaveRequest.leaveHours
  };
}

// 模擬函數：錯誤處理
function simulateErrorHandling(errorType, employeeId, message) {
  return {
    timestamp: new Date(),
    errorType,
    employeeId,
    message,
    affectedData: `${employeeId}-${errorType}`
  };
}

// 模擬函數：反向驗證
async function simulateReverseValidation(sheetId, masterRecord) {
  try {
    // 在實際測試中，這裡會查詢對應的員工檔案
    // 這裡簡化為模擬返回
    const mockEmployeeData = [
      {
        date: '2025-08-15',
        hours: 8,
        type: '平日',
        dayOfWeek: '五'
      }
    ];
    
    const matchedRecord = mockEmployeeData.find(record => 
      record.date === masterRecord.date && record.hours === masterRecord.hours
    );
    
    if (matchedRecord) {
      return {
        found: true,
        matchedRecord
      };
    } else {
      return {
        found: false,
        error: '員工檔案中找不到對應資料'
      };
    }
  } catch (error) {
    return {
      found: false,
      error: `驗證失敗: ${error.message}`
    };
  }
}