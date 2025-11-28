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
    
    // 步驟3: 個人補休表同步
    const syncResult = syncPersonalLeaveRequests();
    const syncedLeaveCount = syncResult.syncedCount;
    errorCount += syncResult.errorCount;
    
    const matchResult = matchLeaveWithOvertime();
    matchedLeaveCount = matchResult.matched;
    errorCount += matchResult.errors;
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    const report = `✅ 檢查完成報告
━━━━━━━━━━━━━━━━━━━━
處理員工數：${processedCount}
新增加班記錄：${newOvertimeCount} 筆
同步補休申請：${syncedLeaveCount} 筆
配對補休記錄：${matchedLeaveCount} 筆
發現錯誤：${errorCount} 筆
執行時間：${duration} 秒

詳細記錄請查看「執行紀錄」工作表`;
    
    Logger.log(report);
    logExecution('完成', '系統', report);
    
    return {
      processedCount,
      newOvertimeCount,
      syncedLeaveCount,
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
      const overtimeHours = parseFloat(row[10]) || 0; // 修正：當日加班時數是第11欄（索引10）
      const note = row[11] || ''; // 修正：備註是第12欄（索引11）
      
      if (date && overtimeHours > 0) {
        let dateObj;
        let formattedDate;
        
        // 更安全的日期處理 - 專為 Google Apps Script 優化
        if (date instanceof Date) {
          // 使用 Google Apps Script 的 Utilities.formatDate 避免時區問題
          try {
            formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            dateObj = date;
          } catch (error) {
            // 如果 Utilities.formatDate 失敗，回退到手動處理
            dateObj = date;
            formattedDate = formatDate(dateObj);
          }
        } else if (typeof date === 'string') {
          // 如果是字符串，嘗試解析
          if (/^\d{4}-\d{2}-\d{2}$/.test(date)) {
            // 已經是正確格式的字符串
            formattedDate = date;
            dateObj = new Date(date + 'T12:00:00'); // 添加中午時間避免時區問題
          } else {
            // 嘗試解析其他格式
            dateObj = new Date(date);
            if (isValidDate(dateObj)) {
              try {
                formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              } catch (error) {
                formattedDate = formatDate(dateObj);
              }
            } else {
              continue; // 跳過無效日期
            }
          }
        } else {
          continue; // 跳過無效日期格式
        }
        
        if (formattedDate && isValidDate(dateObj)) {
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
            date: formattedDate,
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
  
  // 例假日加班處理邏輯
  let remainingHours, status;
  if (record.type === '例假日') {
    // 例假日加班不可補休
    remainingHours = 0;
    status = '例假日-僅發加班費';
  } else {
    // 一般加班可以補休
    remainingHours = record.hours;
    status = '未使用';
  }
  
  const newRow = [
    overtimeId,
    record.employeeId,
    employeeName,
    record.date,
    record.dayOfWeek,
    record.type,
    record.hours,
    0, // 已使用補休
    remainingHours, // 剩餘可補休
    status, // 狀態
    record.sourceMonth,
    '', // 用掉補休編號
    '' // 錯誤提示
  ];
  
  sheet.appendRow(newRow);
  
  // 同步到員工個人加班表
  const employeeInfo = getEmployeeInfo(record.employeeId);
  if (employeeInfo) {
    syncOvertimeToPersonalSheet(employeeInfo.fileId, {
      overtimeId: overtimeId,
      employeeId: record.employeeId,
      employeeName: employeeName,
      date: record.date,
      dayOfWeek: record.dayOfWeek,
      type: record.type,
      hours: record.hours,
      usedHours: 0,
      remainingHours: remainingHours,
      status: status,
      sourceMonth: record.sourceMonth,
      usedLeaveIds: ''
    });
  }
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
        } else {
          // 驗證成功時清空錯誤訊息
          markRecordError(sheet, i + 1, '');
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
          const leaveId = row[0]; // 補休編號
          const result = allocateLeaveToOvertime(employeeId, leaveHours, leaveId);
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

function allocateLeaveToOvertime(employeeId, leaveHours, leaveId = '') {
  const overtimeSheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_SUMMARY);
  const data = overtimeSheet.getDataRange().getValues();
  
  const availableRecords = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === employeeId && 
        parseFloat(row[8]) > 0 && 
        row[9] !== '例假日-僅發加班費') { // 排除例假日加班記錄
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
    
    // 回寫補休編號到加班記錄總表
    if (leaveId) {
      writeLeaveIdToOvertimeRecord(overtimeSheet, record.rowIndex, leaveId);
      
      // 同時回寫到員工個人月份分頁和個人加班表
      const employeeInfo = getEmployeeInfo(employeeId);
      if (employeeInfo) {
        const sourceMonth = data[record.rowIndex - 1][10]; // 資料來源月份
        writeLeaveIdToEmployeeSheet(employeeId, employeeInfo.fileId, record.date, sourceMonth, leaveId);
        
        // 同步更新個人加班表
        const usedLeaveIdsCell = overtimeSheet.getRange(record.rowIndex, 12).getValue() || '';
        updatePersonalOvertimeSheet(
          employeeInfo.fileId,
          record.overtimeId,
          newUsedHours,
          newRemainingHours,
          status,
          usedLeaveIdsCell
        );
      }
    }
    
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
    // 確保日期格式一致的比對
    const existingDate = typeof row[3] === 'string' ? row[3] : formatDate(new Date(row[3]));
    if (row[1] === employeeId && existingDate === date) {
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
    // 確保日期格式一致的比對
    const existingDate = typeof row[3] === 'string' ? row[3] : formatDate(new Date(row[3]));
    if (row[1] === employeeId && existingDate === date) {
      return true;
    }
  }
  
  return false;
}

function verifyRecordInEmployeeFile(fileId, sourceMonth, date) {
  try {
    // 將傳入的 date 統一轉換為字串格式
    let targetDate;
    if (date instanceof Date) {
      try {
        targetDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } catch (error) {
        targetDate = formatDate(date);
      }
    } else if (typeof date === 'string') {
      targetDate = date;
    } else {
      return false; // 無效的日期格式
    }
    
    const sheet = getEmployeeSheet(fileId, sourceMonth);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let rowDate;
      
      // 處理不同的日期格式
      if (row[1] instanceof Date) {
        try {
          rowDate = Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch (error) {
          rowDate = formatDate(row[1]);
        }
      } else if (typeof row[1] === 'string') {
        // 如果是字串，嘗試解析
        if (/^\d{4}-\d{2}-\d{2}$/.test(row[1])) {
          rowDate = row[1]; // 已經是正確格式
        } else {
          const parsedDate = new Date(row[1]);
          if (isValidDate(parsedDate)) {
            try {
              rowDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } catch (error) {
              rowDate = formatDate(parsedDate);
            }
          } else {
            rowDate = row[1]; // 保持原始字串格式
          }
        }
      } else {
        continue; // 跳過無效的日期格式
      }
      
      if (rowDate === targetDate && parseFloat(row[10]) > 0) { // 加班時數在第11欄（索引10）
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

function formatDateSafe(date) {
  try {
    // 處理時區問題：使用 UTC 方法或本地偏移
    if (date instanceof Date) {
      // 檢查是否為有效的日期對象
      if (isNaN(date.getTime())) {
        return null;
      }
      
      // 如果是 Google Sheets 的日期（可能有時區偏移），使用 UTC
      const timezoneOffset = date.getTimezoneOffset();
      const adjustedDate = new Date(date.getTime() + (timezoneOffset * 60000));
      
      const year = adjustedDate.getFullYear();
      const month = String(adjustedDate.getMonth() + 1).padStart(2, '0');
      const day = String(adjustedDate.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    return null;
  } catch (error) {
    return null;
  }
}

function getDayOfWeek(date) {
  const days = ['日', '一', '二', '三', '四', '五', '六'];
  return days[date.getDay()];
}

// ===== 個人補休表同步功能 =====

function syncPersonalLeaveRequests() {
  try {
    const startTime = new Date();
    logExecution('開始', '系統', '開始執行個人補休表同步');
    
    let syncedCount = 0;
    let errorCount = 0;
    
    const employees = getActiveEmployees();
    
    for (const employee of employees) {
      try {
        const leaveRequests = scanPersonalLeaveSheet(employee.fileId);
        
        for (const request of leaveRequests) {
          try {
            const leaveId = generateLeaveId(request.applicationDate, employee.id);
            const leaveRequest = {
              leaveId: leaveId,
              employeeId: employee.id,
              employeeName: employee.name,
              applicationDate: request.applicationDate,
              leaveDate: request.leaveDate,
              startTime: request.startTime,
              endTime: request.endTime,
              hours: request.hours,
              note: request.note
            };
            
            const result = addLeaveRequestToMaster(leaveRequest);
            if (result.success) {
              writeLeaveIdToPersonalSheet(employee.fileId, request.rowIndex, leaveId);
              syncedCount++;
            } else {
              errorCount++;
              logError('補休申請同步失敗', employee.id, result.error);
            }
          } catch (error) {
            errorCount++;
            logError('個人補休處理失敗', employee.id, error.message);
          }
        }
        
        logExecution('處理', employee.id, `個人補休表同步完成 - 新增 ${leaveRequests.length} 筆申請`);
      } catch (error) {
        errorCount++;
        logError('員工補休表存取失敗', employee.id, error.message);
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    const report = `個人補休表同步完成：同步 ${syncedCount} 筆，錯誤 ${errorCount} 筆，耗時 ${duration} 秒`;
    logExecution('完成', '系統', report);
    
    return { syncedCount, errorCount, duration };
  } catch (error) {
    logError('個人補休表同步失敗', '系統', error.message);
    throw error;
  }
}

function scanPersonalLeaveSheet(fileId) {
  try {
    const sheet = getEmployeeSheet(fileId, CONFIG.SHEETS.PERSONAL_LEAVE);
    const data = sheet.getDataRange().getValues();
    const leaveRequests = [];
    
    Logger.log(`掃描個人補休表，共 ${data.length - 1} 筆資料`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const leaveId = row[0]; // 補休編號
      const employeeId = row[1]; // 員工編號
      const employeeName = row[2]; // 姓名
      const applicationDate = row[3]; // 申請日期
      const leaveDate = row[4]; // 補休日期
      const startTime = row[5]; // 開始時間
      const endTime = row[6]; // 結束時間
      const hours = parseFloat(row[7]) || 0; // 使用時數
      const overtimeIds = row[8] || ''; // 對應加班編號
      const note = row[9] || ''; // 備註
      
      Logger.log(`第${i+1}列: 補休編號=[${leaveId}], 員工編號=[${employeeId}], 申請日期=[${applicationDate}], 補休日期=[${leaveDate}], 時數=[${hours}]`);
      
      // 只處理尚未同步的記錄（補休編號為空）
      if (!leaveId && applicationDate && leaveDate && hours > 0) {
        let appDateFormatted, lvDateFormatted;
        let appDateObj, lvDateObj;
        
        // 安全的申請日期處理
        if (applicationDate instanceof Date) {
          try {
            appDateFormatted = Utilities.formatDate(applicationDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            appDateObj = applicationDate;
          } catch (error) {
            appDateFormatted = formatDate(applicationDate);
            appDateObj = applicationDate;
          }
        } else if (typeof applicationDate === 'string') {
          if (/^\d{4}-\d{2}-\d{2}$/.test(applicationDate)) {
            appDateFormatted = applicationDate;
            appDateObj = new Date(applicationDate + 'T12:00:00');
          } else {
            appDateObj = new Date(applicationDate);
            if (isValidDate(appDateObj)) {
              try {
                appDateFormatted = Utilities.formatDate(appDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              } catch (error) {
                appDateFormatted = formatDate(appDateObj);
              }
            } else {
              appDateFormatted = null;
            }
          }
        } else {
          appDateFormatted = null;
        }
        
        // 安全的補休日期處理
        if (leaveDate instanceof Date) {
          try {
            lvDateFormatted = Utilities.formatDate(leaveDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            lvDateObj = leaveDate;
          } catch (error) {
            lvDateFormatted = formatDate(leaveDate);
            lvDateObj = leaveDate;
          }
        } else if (typeof leaveDate === 'string') {
          if (/^\d{4}-\d{2}-\d{2}$/.test(leaveDate)) {
            lvDateFormatted = leaveDate;
            lvDateObj = new Date(leaveDate + 'T12:00:00');
          } else {
            lvDateObj = new Date(leaveDate);
            if (isValidDate(lvDateObj)) {
              try {
                lvDateFormatted = Utilities.formatDate(lvDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              } catch (error) {
                lvDateFormatted = formatDate(lvDateObj);
              }
            } else {
              lvDateFormatted = null;
            }
          }
        } else {
          lvDateFormatted = null;
        }
        
        Logger.log(`檢查日期有效性: 申請日期=${appDateFormatted}, 補休日期=${lvDateFormatted}`);
        
        if (appDateFormatted && lvDateFormatted && isValidDate(appDateObj) && isValidDate(lvDateObj)) {
          leaveRequests.push({
            rowIndex: i + 1,
            employeeId: employeeId,
            employeeName: employeeName,
            applicationDate: appDateFormatted,
            leaveDate: lvDateFormatted,
            startTime: startTime,
            endTime: endTime,
            hours: hours,
            note: note
          });
          Logger.log(`✅ 加入處理清單: 第${i+1}列`);
        } else {
          Logger.log(`❌ 日期格式無效: 第${i+1}列`);
        }
      } else {
        Logger.log(`❌ 跳過第${i+1}列 - 補休編號=${leaveId ? '已有' : '空'}, 申請日期=${applicationDate ? '有' : '無'}, 補休日期=${leaveDate ? '有' : '無'}, 時數=${hours}`);
      }
    }
    
    Logger.log(`掃描完成，找到 ${leaveRequests.length} 筆待處理記錄`);
    return leaveRequests;
  } catch (error) {
    // 如果找不到個人補休表分頁，返回空陣列
    if (error.message.includes('找不到工作表')) {
      Logger.log(`員工檔案中沒有「${CONFIG.SHEETS.PERSONAL_LEAVE}」工作表，跳過處理`);
      return [];
    }
    throw new Error(`掃描個人補休表失敗: ${error.message}`);
  }
}

function generateLeaveId(date, employeeId) {
  const dateStr = date.replace(/-/g, '');
  const existingIds = getExistingLeaveIds(date, employeeId);
  const sequence = existingIds.length + 1;
  return `LV-${dateStr}-${employeeId}-${sequence}`;
}

function getExistingLeaveIds(date, employeeId) {
  try {
    const sheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_DETAILS);
    const data = sheet.getDataRange().getValues();
    const ids = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === employeeId) {
        const leaveId = row[0];
        if (leaveId && leaveId.includes(date.replace(/-/g, ''))) {
          ids.push(leaveId);
        }
      }
    }
    
    return ids;
  } catch (error) {
    return [];
  }
}

function addLeaveRequestToMaster(leaveRequest) {
  try {
    // 先進行自動配對
    const allocationResult = allocateLeaveToOvertime(leaveRequest.employeeId, leaveRequest.hours, leaveRequest.leaveId);
    
    const sheet = getMasterSheet(CONFIG.SHEETS.OVERTIME_DETAILS);
    
    let overtimeIds = '';
    let error = '';
    
    if (allocationResult.success) {
      overtimeIds = allocationResult.overtimeIds.join(',');
    } else {
      error = allocationResult.error;
    }
    
    const newRow = [
      leaveRequest.leaveId,
      leaveRequest.employeeId,
      leaveRequest.employeeName,
      leaveRequest.applicationDate,
      leaveRequest.leaveDate,
      leaveRequest.startTime || '', // 開始時間
      leaveRequest.endTime || '', // 結束時間
      leaveRequest.hours,
      overtimeIds, // 對應加班編號
      leaveRequest.note,
      false, // 行政組查閱打勾（核選方框）
      error // 錯誤提示
    ];
    
    sheet.appendRow(newRow);
    
    return allocationResult;
  } catch (error) {
    throw new Error(`新增補休申請失敗: ${error.message}`);
  }
}

function writeLeaveIdToPersonalSheet(fileId, rowIndex, leaveId) {
  try {
    const sheet = getEmployeeSheet(fileId, CONFIG.SHEETS.PERSONAL_LEAVE);
    sheet.getRange(rowIndex, 1).setValue(leaveId); // 第1欄是補休編號
    return true;
  } catch (error) {
    logError('回寫補休編號失敗', fileId, error.message);
    return false;
  }
}

function writeLeaveIdToOvertimeRecord(overtimeSheet, rowIndex, leaveId) {
  try {
    const currentCell = overtimeSheet.getRange(rowIndex, 12); // 第12欄是用掉補休編號
    const currentValue = currentCell.getValue() || '';
    
    let newValue;
    if (currentValue === '') {
      newValue = leaveId;
    } else {
      // 如果已有補休編號，用逗號分隔
      const existingIds = currentValue.split(',');
      if (!existingIds.includes(leaveId)) {
        existingIds.push(leaveId);
        newValue = existingIds.join(',');
      } else {
        newValue = currentValue; // 不重複新增
      }
    }
    
    currentCell.setValue(newValue);
    return true;
  } catch (error) {
    logError('回寫補休編號到加班記錄失敗', 'system', error.message);
    return false;
  }
}

function writeLeaveIdToEmployeeSheet(employeeId, fileId, date, sourceMonth, leaveId) {
  try {
    const sheet = getEmployeeSheet(fileId, sourceMonth);
    const data = sheet.getDataRange().getValues();
    
    // 找到對應日期的列
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let rowDate;
      
      if (row[1] instanceof Date) {
        try {
          rowDate = Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch (error) {
          rowDate = formatDate(row[1]);
        }
      } else if (typeof row[1] === 'string') {
        if (/^\d{4}-\d{2}-\d{2}$/.test(row[1])) {
          rowDate = row[1]; // 已經是正確格式
        } else {
          const parsedDate = new Date(row[1]);
          if (isValidDate(parsedDate)) {
            try {
              rowDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } catch (error) {
              rowDate = formatDate(parsedDate);
            }
          } else {
            rowDate = row[1];
          }
        }
      } else {
        continue;
      }
      
      if (rowDate === date) {
        // 在備註欄（第13欄）加入補休編號
        const noteCell = sheet.getRange(i + 1, 13);
        const currentNote = noteCell.getValue() || '';
        
        let newNote;
        if (currentNote === '') {
          newNote = `補休編號:${leaveId}`;
        } else {
          if (!currentNote.includes(leaveId)) {
            newNote = `${currentNote}; 補休編號:${leaveId}`;
          } else {
            newNote = currentNote; // 不重複新增
          }
        }
        
        noteCell.setValue(newNote);
        return true;
      }
    }
    
    return false; // 找不到對應日期
  } catch (error) {
    logError('回寫補休編號到員工表格失敗', employeeId, error.message);
    return false;
  }
}

// ===== 個人加班表同步 =====

function syncOvertimeToPersonalSheet(fileId, overtimeRecord) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    let personalOvertimeSheet = ss.getSheetByName(CONFIG.SHEETS.PERSONAL_OVERTIME);
    
    // 如果個人加班表不存在，則建立它
    if (!personalOvertimeSheet) {
      personalOvertimeSheet = ss.insertSheet(CONFIG.SHEETS.PERSONAL_OVERTIME);
      
      // 建立標題列
      const headers = [
        '加班編號',
        '員工編號',
        '姓名',
        '日期',
        '星期',
        '加班類型',
        '加班時數',
        '已使用補休',
        '剩餘可補休',
        '狀態',
        '資料來源月份',
        '用掉補休編號'
      ];
      personalOvertimeSheet.appendRow(headers);
      
      // 設定標題列格式
      const headerRange = personalOvertimeSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
      
      Logger.log(`✅ 已為員工 ${overtimeRecord.employeeId} 建立「${CONFIG.SHEETS.PERSONAL_OVERTIME}」工作表`);
    }
    
    // 檢查該加班編號是否已存在
    const data = personalOvertimeSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === overtimeRecord.overtimeId) {
        Logger.log(`個人加班表已存在記錄: ${overtimeRecord.overtimeId}`);
        return; // 已存在，不重複新增
      }
    }
    
    // 新增記錄到個人加班表
    const newRow = [
      overtimeRecord.overtimeId,
      overtimeRecord.employeeId,
      overtimeRecord.employeeName,
      overtimeRecord.date,
      overtimeRecord.dayOfWeek,
      overtimeRecord.type,
      overtimeRecord.hours,
      overtimeRecord.usedHours,
      overtimeRecord.remainingHours,
      overtimeRecord.status,
      overtimeRecord.sourceMonth,
      overtimeRecord.usedLeaveIds
    ];
    
    personalOvertimeSheet.appendRow(newRow);
    Logger.log(`✅ 已同步加班記錄到個人加班表: ${overtimeRecord.overtimeId}`);
    
    return true;
  } catch (error) {
    logError('同步到個人加班表失敗', overtimeRecord.employeeId, error.message);
    return false;
  }
}

function updatePersonalOvertimeSheet(fileId, overtimeId, usedHours, remainingHours, status, usedLeaveIds) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const personalOvertimeSheet = ss.getSheetByName(CONFIG.SHEETS.PERSONAL_OVERTIME);
    
    if (!personalOvertimeSheet) {
      Logger.log(`找不到「${CONFIG.SHEETS.PERSONAL_OVERTIME}」工作表，跳過更新`);
      return false;
    }
    
    const data = personalOvertimeSheet.getDataRange().getValues();
    
    // 找到對應的加班記錄並更新
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === overtimeId) {
        personalOvertimeSheet.getRange(i + 1, 8).setValue(usedHours); // 已使用補休
        personalOvertimeSheet.getRange(i + 1, 9).setValue(remainingHours); // 剩餘可補休
        personalOvertimeSheet.getRange(i + 1, 10).setValue(status); // 狀態
        
        // 更新用掉補休編號
        if (usedLeaveIds) {
          const currentCell = personalOvertimeSheet.getRange(i + 1, 12);
          const currentValue = currentCell.getValue() || '';
          
          if (currentValue === '') {
            currentCell.setValue(usedLeaveIds);
          } else {
            // 合併補休編號（避免重複）
            const existingIds = currentValue.split(',').map(id => id.trim());
            const newIds = usedLeaveIds.split(',').map(id => id.trim());
            
            newIds.forEach(id => {
              if (!existingIds.includes(id)) {
                existingIds.push(id);
              }
            });
            
            currentCell.setValue(existingIds.join(','));
          }
        }
        
        Logger.log(`✅ 已更新個人加班表記錄: ${overtimeId}`);
        return true;
      }
    }
    
    Logger.log(`在個人加班表中找不到記錄: ${overtimeId}`);
    return false;
  } catch (error) {
    logError('更新個人加班表失敗', fileId, error.message);
    return false;
  }
}

// ===== 測試與驗證函數 =====

function debugPersonalLeaveSheet() {
  const fileId = '1g82f-yigavTtmL3YABxINFaXHerDppve9Ol0CM7aHVo';
  
  try {
    const ss = SpreadsheetApp.openById(fileId);
    const sheets = ss.getSheets();
    
    Logger.log(`試算表共有 ${sheets.length} 個工作表:`);
    sheets.forEach((sheet, index) => {
      Logger.log(`${index + 1}. ${sheet.getName()}`);
    });
    
    // 尋找個人補休表
    const personalLeaveSheet = ss.getSheetByName(CONFIG.SHEETS.PERSONAL_LEAVE);
    if (!personalLeaveSheet) {
      Logger.log(`❌ 找不到「${CONFIG.SHEETS.PERSONAL_LEAVE}」工作表`);
      return;
    }
    
    Logger.log(`✅ 找到「${CONFIG.SHEETS.PERSONAL_LEAVE}」工作表`);
    
    // 讀取前 3 列來檢查結構 (擴展到15欄)
    const data = personalLeaveSheet.getRange(1, 1, 3, 15).getValues();
    
    Logger.log('=== 個人補休表結構 ===');
    Logger.log('標題列 (第1列):');
    data[0].forEach((cell, index) => {
      Logger.log(`欄位 ${index}: [${cell}]`);
    });
    
    if (data.length > 1) {
      Logger.log('資料列 (第2列):');
      data[1].forEach((cell, index) => {
        Logger.log(`欄位 ${index}: [${cell}]`);
      });
    }
    
    if (data.length > 2) {
      Logger.log('資料列 (第3列):');
      data[2].forEach((cell, index) => {
        Logger.log(`欄位 ${index}: [${cell}]`);
      });
    }
    
    Logger.log(`總共有 ${personalLeaveSheet.getLastRow()} 列資料`);
    
  } catch (error) {
    Logger.log(`❌ 檢查失敗: ${error.message}`);
  }
}

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