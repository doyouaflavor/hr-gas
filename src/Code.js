function testConnection() {
  Logger.log('測試連線...');
  
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
  const employeeSheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
  
  if (employeeSheet) {
    Logger.log('✅ 成功連接到主控制台');
    Logger.log('員工數量：' + (employeeSheet.getLastRow() - 1));
  } else {
    Logger.log('❌ 找不到員工清單工作表');
  }
}

// ===== 主要流程 =====

function checkAllEmployees() {
  try {
    const startTime = new Date();
    logExecution('開始', '系統', '開始執行全員檢查');
    
    let processedCount = 0;
    let newOvertimeCount = 0;
    let matchedLeaveCount = 0;
    let errorCount = 0;
    
    const employees = getActiveEmployees();
    
    for (const employee of employees) {
      try {
        const result = checkEmployeeData(employee.id, employee.fileId);
        newOvertimeCount += result.newRecords;
        processedCount++;
        
        logExecution('處理', employee.id, `處理完成 - 新增 ${result.newRecords} 筆加班記錄`);
      } catch (error) {
        errorCount++;
        logError('員工處理失敗', employee.id, error.message);
      }
    }
    
    const validationErrors = validateOvertimeRecords();
    errorCount += validationErrors;
    
    const matchResult = matchLeaveWithOvertime();
    matchedLeaveCount = matchResult.matched;
    errorCount += matchResult.errors;
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    const report = `✅ 檢查完成報告
━━━━━━━━━━━━━━━━━━━━
處理員工數：${processedCount}
新增加班記錄：${newOvertimeCount} 筆
配對補休記錄：${matchedLeaveCount} 筆
發現錯誤：${errorCount} 筆
執行時間：${duration} 秒

詳細記錄請查看「執行紀錄」工作表`;
    
    Logger.log(report);
    logExecution('完成', '系統', report);
    
    return {
      processedCount,
      newOvertimeCount,
      matchedLeaveCount,
      errorCount,
      duration
    };
    
  } catch (error) {
    logError('系統錯誤', '系統', error.message);
    throw error;
  }
}

// ===== 員工資料檢查 =====

function checkEmployeeData(employeeId, fileId) {
  try {
    const sheets = scanMonthlySheets(fileId);
    let newRecordCount = 0;
    
    for (const sheetInfo of sheets) {
      try {
        const overtimeData = extractOvertimeData(sheetInfo.sheet, sheetInfo.name, employeeId);
        
        for (const record of overtimeData) {
          if (!overtimeRecordExists(employeeId, record.date)) {
            addOvertimeRecord(record);
            newRecordCount++;
          }
        }
      } catch (error) {
        logError('月份分頁處理失敗', employeeId, `${sheetInfo.name}: ${error.message}`);
      }
    }
    
    return { newRecords: newRecordCount };
  } catch (error) {
    logError('員工資料檢查失敗', employeeId, error.message);
    throw error;
  }
}

function scanMonthlySheets(fileId) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const sheets = ss.getSheets();
    const monthlySheets = [];
    
    const monthPattern = /^(\d{1,2})月$/;
    
    for (const sheet of sheets) {
      const name = sheet.getName();
      if (monthPattern.test(name)) {
        monthlySheets.push({
          sheet: sheet,
          name: name,
          month: parseInt(name.match(monthPattern)[1])
        });
      }
    }
    
    monthlySheets.sort((a, b) => a.month - b.month);
    return monthlySheets;
  } catch (error) {
    throw new Error(`無法存取員工檔案: ${error.message}`);
  }
}

function extractOvertimeData(sheet, monthName, employeeId) {
  try {
    const data = sheet.getDataRange().getValues();
    const overtimeRecords = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[1];
      const dayOfWeek = row[2];
      const dayType = row[3] || '';
      const overtimeHours = parseFloat(row[11]) || 0;
      const note = row[12] || '';
      
      if (date && overtimeHours > 0) {
        const dateObj = new Date(date);
        if (isValidDate(dateObj)) {
          let overtimeType = '';
          if (dayType.includes('例假日')) {
            overtimeType = '例假日';
          } else if (dayType.includes('休息日') || dayOfWeek === '6' || dayOfWeek === '7') {
            overtimeType = '假日';
          } else if (dayType.includes('上班日加班')) {
            overtimeType = '上班日加班';
          } else if (dayType.includes('上班日')) {
            overtimeType = '上班日';
          } else {
            overtimeType = '平日';
          }
          
          overtimeRecords.push({
            employeeId: employeeId,
            date: formatDate(dateObj),
            hours: overtimeHours,
            type: overtimeType,
            note: note,
            sourceMonth: monthName,
            dayOfWeek: typeof dayOfWeek === 'number' ? getDayOfWeek(dateObj) : dayOfWeek
          });
        }
      }
    }
    
    return overtimeRecords;
  } catch (error) {
    throw new Error(`分頁資料提取失敗: ${error.message}`);
  }
}

function addOvertimeRecord(record) {
  const sheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
  const employeeName = getEmployeeName(record.employeeId);
  const overtimeId = generateOvertimeId(record.date, record.employeeId);
  
  const newRow = [
    overtimeId,
    record.employeeId,
    employeeName,
    record.date,
    record.dayOfWeek,
    record.type,
    record.hours,
    0,
    record.hours,
    '未使用',
    record.sourceMonth,
    '',
    ''
  ];
  
  sheet.appendRow(newRow);
}

// ===== 反向驗證 =====

function validateOvertimeRecords() {
  try {
    const sheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const data = sheet.getDataRange().getValues();
    let errorCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employeeId = row[1];
      const date = row[3];
      const sourceMonth = row[10];
      
      try {
        const employee = getEmployeeInfo(employeeId);
        if (!employee) {
          markRecordError(sheet, i + 1, '找不到員工資料');
          errorCount++;
          continue;
        }
        
        const found = verifyRecordInEmployeeFile(employee.fileId, sourceMonth, date);
        if (!found) {
          markRecordError(sheet, i + 1, '員工檔案中找不到對應資料');
          errorCount++;
        }
      } catch (error) {
        markRecordError(sheet, i + 1, `驗證失敗: ${error.message}`);
        errorCount++;
      }
    }
    
    return errorCount;
  } catch (error) {
    logError('反向驗證失敗', '系統', error.message);
    return 0;
  }
}

// ===== 補休配對 =====

function matchLeaveWithOvertime() {
  try {
    const leaveSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_DETAILS);
    const leaveData = leaveSheet.getDataRange().getValues();
    
    let matchedCount = 0;
    let errorCount = 0;
    
    for (let i = 1; i < leaveData.length; i++) {
      const row = leaveData[i];
      const employeeId = row[1];
      const leaveHours = parseFloat(row[5]) || 0;
      const existingOvertimeId = row[6];
      
      if (!existingOvertimeId || existingOvertimeId === '') {
        try {
          const result = allocateLeaveToOvertime(employeeId, leaveHours);
          if (result.success) {
            updateLeaveRecord(leaveSheet, i + 1, result.overtimeIds.join(','));
            matchedCount++;
          } else {
            markLeaveError(leaveSheet, i + 1, result.error);
            errorCount++;
          }
        } catch (error) {
          markLeaveError(leaveSheet, i + 1, error.message);
          errorCount++;
        }
      }
    }
    
    return { matched: matchedCount, errors: errorCount };
  } catch (error) {
    logError('補休配對失敗', '系統', error.message);
    return { matched: 0, errors: 1 };
  }
}

function allocateLeaveToOvertime(employeeId, leaveHours) {
  const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
  const data = overtimeSheet.getDataRange().getValues();
  
  const availableRecords = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === employeeId && parseFloat(row[8]) > 0) {
      availableRecords.push({
        rowIndex: i + 1,
        overtimeId: row[0],
        date: row[3],
        totalHours: parseFloat(row[6]),
        usedHours: parseFloat(row[7]),
        remainingHours: parseFloat(row[8])
      });
    }
  }
  
  availableRecords.sort((a, b) => new Date(a.date) - new Date(b.date));
  
  let remainingLeaveHours = leaveHours;
  const usedOvertimeIds = [];
  
  for (const record of availableRecords) {
    if (remainingLeaveHours <= 0) break;
    
    const allocatedHours = Math.min(remainingLeaveHours, record.remainingHours);
    const newUsedHours = record.usedHours + allocatedHours;
    const newRemainingHours = record.remainingHours - allocatedHours;
    
    let status = '未使用';
    if (newRemainingHours === 0) {
      status = '已全數使用';
    } else if (newUsedHours > 0) {
      status = '部分使用';
    }
    
    overtimeSheet.getRange(record.rowIndex, 8).setValue(newUsedHours);
    overtimeSheet.getRange(record.rowIndex, 9).setValue(newRemainingHours);
    overtimeSheet.getRange(record.rowIndex, 10).setValue(status);
    
    usedOvertimeIds.push(record.overtimeId);
    remainingLeaveHours -= allocatedHours;
  }
  
  if (remainingLeaveHours > 0) {
    return {
      success: false,
      error: `補休時數超過可用加班時數 (超過 ${remainingLeaveHours} 小時)`
    };
  }
  
  return {
    success: true,
    overtimeIds: usedOvertimeIds
  };
}

// ===== 工具函數 =====

function getMasterSheet(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`找不到工作表: ${sheetName}`);
  }
  return sheet;
}

function getEmployeeSheet(fileId, sheetName) {
  const ss = SpreadsheetApp.openById(fileId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`找不到工作表: ${sheetName}`);
  }
  return sheet;
}

function generateOvertimeId(date, employeeId) {
  const dateStr = date.replace(/-/g, '');
  const existingIds = getExistingOvertimeIds(date, employeeId);
  const sequence = existingIds.length + 1;
  return `OT-${dateStr}-${employeeId}-${sequence}`;
}

function getExistingOvertimeIds(date, employeeId) {
  const sheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
  const data = sheet.getDataRange().getValues();
  const ids = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === employeeId && row[3] === date) {
      ids.push(row[0]);
    }
  }
  
  return ids;
}

function getActiveEmployees() {
  const sheet = getMasterSheet(CONFIG.SHEETS.EMPLOYEES);
  const data = sheet.getDataRange().getValues();
  const employees = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[4] === '啟用') {
      employees.push({
        id: row[0],
        name: row[1],
        fileId: row[2],
        status: row[4]
      });
    }
  }
  
  return employees;
}

function getEmployeeInfo(employeeId) {
  const sheet = getMasterSheet(CONFIG.SHEETS.EMPLOYEES);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === employeeId) {
      return {
        id: row[0],
        name: row[1],
        fileId: row[2],
        status: row[4]
      };
    }
  }
  
  return null;
}

function getEmployeeName(employeeId) {
  const employee = getEmployeeInfo(employeeId);
  return employee ? employee.name : '未知員工';
}

function overtimeRecordExists(employeeId, date) {
  const sheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === employeeId && row[3] === date) {
      return true;
    }
  }
  
  return false;
}

function verifyRecordInEmployeeFile(fileId, sourceMonth, date) {
  try {
    const sheet = getEmployeeSheet(fileId, sourceMonth);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDate(new Date(row[0]));
      if (rowDate === date && parseFloat(row[4]) > 0) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    return false;
  }
}

function markRecordError(sheet, rowIndex, errorMessage) {
  sheet.getRange(rowIndex, 13).setValue(errorMessage);
}

function markLeaveError(sheet, rowIndex, errorMessage) {
  sheet.getRange(rowIndex, 10).setValue(errorMessage);
}

function updateLeaveRecord(sheet, rowIndex, overtimeIds) {
  sheet.getRange(rowIndex, 7).setValue(overtimeIds);
}

function logError(type, employeeId, message) {
  try {
    const sheet = getMasterSheet(CONFIG.SHEETS.EXECUTION_LOG);
    const timestamp = new Date();
    const newRow = [timestamp, 'ERROR', type, employeeId, message, ''];
    sheet.appendRow(newRow);
    Logger.log(`ERROR: ${type} - ${employeeId}: ${message}`);
  } catch (error) {
    Logger.log(`記錄錯誤失敗: ${error.message}`);
  }
}

function logExecution(type, employeeId, message) {
  try {
    const sheet = getMasterSheet(CONFIG.SHEETS.EXECUTION_LOG);
    const timestamp = new Date();
    const newRow = [timestamp, 'INFO', type, employeeId, message, ''];
    sheet.appendRow(newRow);
  } catch (error) {
    Logger.log(`記錄執行日誌失敗: ${error.message}`);
  }
}

function isValidDate(date) {
  return date instanceof Date && !isNaN(date);
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getDayOfWeek(date) {
  const days = ['日', '一', '二', '三', '四', '五', '六'];
  return days[date.getDay()];
}

// ===== 測試與驗證函數 =====

function validateSystemSetup() {
  const results = {
    success: true,
    errors: [],
    warnings: []
  };
  
  try {
    const masterSS = SpreadsheetApp.openById(CONFIG.MASTER_SHEET_ID);
    Logger.log('✅ 主控台連接成功');
    
    const requiredSheets = Object.values(CONFIG.SHEETS);
    for (const sheetName of requiredSheets) {
      const sheet = masterSS.getSheetByName(sheetName);
      if (!sheet) {
        results.errors.push(`缺少必要工作表: ${sheetName}`);
        results.success = false;
      } else {
        Logger.log(`✅ 工作表 "${sheetName}" 存在`);
      }
    }
    
    const employees = getActiveEmployees();
    Logger.log(`✅ 找到 ${employees.length} 位啟用員工`);
    
    let accessibleCount = 0;
    for (const employee of employees) {
      try {
        SpreadsheetApp.openById(employee.fileId);
        accessibleCount++;
      } catch (error) {
        results.warnings.push(`無法存取員工 ${employee.id} 的檔案: ${error.message}`);
      }
    }
    
    Logger.log(`✅ 可存取 ${accessibleCount}/${employees.length} 個員工檔案`);
    
  } catch (error) {
    results.errors.push(`系統設定驗證失敗: ${error.message}`);
    results.success = false;
  }
  
  return results;
}

function runSystemTest() {
  Logger.log('🧪 開始系統測試...');
  
  const validation = validateSystemSetup();
  if (!validation.success) {
    Logger.log('❌ 系統設定驗證失敗');
    validation.errors.forEach(error => Logger.log(`ERROR: ${error}`));
    return false;
  }
  
  if (validation.warnings.length > 0) {
    Logger.log('⚠️ 發現警告:');
    validation.warnings.forEach(warning => Logger.log(`WARNING: ${warning}`));
  }
  
  try {
    Logger.log('📊 執行小規模測試...');
    
    const employees = getActiveEmployees().slice(0, 2);
    
    for (const employee of employees) {
      const result = checkEmployeeData(employee.id, employee.fileId);
      Logger.log(`✅ 員工 ${employee.id} 測試完成 - 新增 ${result.newRecords} 筆記錄`);
    }
    
    Logger.log('🔍 測試反向驗證...');
    const validationErrors = validateOvertimeRecords();
    Logger.log(`✅ 反向驗證完成 - 發現 ${validationErrors} 個錯誤`);
    
    Logger.log('🔗 測試補休配對...');
    const matchResult = matchLeaveWithOvertime();
    Logger.log(`✅ 補休配對完成 - 配對 ${matchResult.matched} 筆，錯誤 ${matchResult.errors} 筆`);
    
    Logger.log('✅ 系統測試完成');
    return true;
    
  } catch (error) {
    Logger.log(`❌ 系統測試失敗: ${error.message}`);
    return false;
  }
}

function getSystemStatus() {
  try {
    const employees = getActiveEmployees();
    const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
    const leaveSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_DETAILS);
    
    const overtimeCount = overtimeSheet.getLastRow() - 1;
    const leaveCount = leaveSheet.getLastRow() - 1;
    
    const unmatchedLeave = leaveSheet.getDataRange().getValues()
      .slice(1)
      .filter(row => !row[6] || row[6] === '').length;
    
    const status = {
      employees: employees.length,
      overtimeRecords: overtimeCount,
      leaveRecords: leaveCount,
      unmatchedLeave: unmatchedLeave,
      lastUpdate: new Date()
    };
    
    Logger.log('📊 系統狀態:');
    Logger.log(`員工數: ${status.employees}`);
    Logger.log(`加班記錄: ${status.overtimeRecords}`);
    Logger.log(`補休記錄: ${status.leaveRecords}`);
    Logger.log(`未配對補休: ${status.unmatchedLeave}`);
    
    return status;
  } catch (error) {
    Logger.log(`❌ 獲取系統狀態失敗: ${error.message}`);
    return null;
  }
}