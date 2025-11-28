// ===== 選單功能 =====

/**
 * 創建自定義選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('加班時數計算')
    .addItem('計算當前試算表的加班時數', 'runOvertimeCalculation')
    .addItem('測試加班計算功能', 'testConnection')
    .addItem('查看執行紀錄', 'showExecutionLog')
    .addToUi();
}

/**
 * 主要功能：計算控制中樞表中已啟用員工的加班時數
 */
function runOvertimeCalculation() {
  const startTime = new Date();
  const CONTROL_CENTER_ID = '1fTQ3AZ93yP_q7oCncMASozScIJ36NlJwgc3vplr0nJI';
  
  try {
    // 記錄開始
    logExecution('開始', '系統', '開始計算所有啟用員工加班時數', '控制中樞表', '讀取員工清單並批量處理');
    
    // 讀取控制中樞表的員工清單
    const activeEmployees = getActiveEmployeesFromControlCenter(CONTROL_CENTER_ID);
    
    if (activeEmployees.length === 0) {
      const noEmployeeMessage = '⚠️ 沒有找到已啟用的員工';
      logExecution('警告', '系統', '無已啟用員工', '控制中樞表', '員工清單中沒有狀態為「啟用」的員工');
      SpreadsheetApp.getUi().alert('警告', noEmployeeMessage, SpreadsheetApp.getUi().AlertType.WARNING);
      return;
    }
    
    let allOvertimeRecords = [];
    let successCount = 0;
    let errorDetails = [];
    
    // 逐一處理每個已啟用的員工
    for (const employee of activeEmployees) {
      try {
        logExecution('處理', '系統', `開始處理員工: ${employee.name}`, '個人試算表', `ID: ${employee.fileId}`);
        
        const overtimeRecords = calculateOvertimeForEmployee(employee.fileId);
        
        // 為每筆記錄加上員工資訊
        const recordsWithEmployee = overtimeRecords.map(record => ({
          ...record,
          employeeId: employee.id,
          employeeName: employee.name
        }));
        
        allOvertimeRecords.push(...recordsWithEmployee);
        successCount++;
        
        logExecution('完成', '系統', `完成處理員工: ${employee.name}`, '個人試算表', 
                    `找到 ${overtimeRecords.length} 筆加班記錄`);
                    
      } catch (error) {
        const errorMsg = `處理員工 ${employee.name} (${employee.id}) 失敗: ${error.message}`;
        errorDetails.push(errorMsg);
        logExecution('錯誤', '系統', `處理員工失敗: ${employee.name}`, '個人試算表', error.message);
        console.log(errorMsg);
      }
    }
    
    // 為每位員工寫入加班記錄到各自的「加班費紀錄總表」
    let totalRecordsWritten = 0;
    for (const employee of activeEmployees) {
      const employeeRecords = allOvertimeRecords.filter(record => record.employeeId === employee.id);
      if (employeeRecords.length > 0) {
        const recordCount = writeOvertimeRecordsToEmployeeSheet(employeeRecords, employee.fileId);
        totalRecordsWritten += recordCount;
      }
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    let successMessage = `✅ 批量加班時數計算完成！
━━━━━━━━━━━━━━━━━━━━
處理員工數: ${activeEmployees.length} 人
成功處理: ${successCount} 人
共計算出: ${allOvertimeRecords.length} 筆加班記錄
成功寫入: ${totalRecordsWritten} 筆到各員工試算表
執行時間: ${duration.toFixed(2)} 秒`;
    
    if (errorDetails.length > 0) {
      successMessage += `\n\n❌ 處理失敗: ${errorDetails.length} 人\n${errorDetails.slice(0, 3).join('\n')}`;
      if (errorDetails.length > 3) {
        successMessage += `\n... 還有 ${errorDetails.length - 3} 個錯誤，請查看執行紀錄`;
      }
    }
    
    // 記錄完成
    logExecution('完成', '系統', '批量加班時數計算完成', '所有員工試算表', 
                `處理: ${activeEmployees.length}人, 成功: ${successCount}人, 記錄: ${allOvertimeRecords.length}筆, 寫入: ${totalRecordsWritten}筆, 耗時: ${duration.toFixed(2)}秒`);
    
    // 顯示結果給用戶
    SpreadsheetApp.getUi().alert('批量計算完成', successMessage, SpreadsheetApp.getUi().AlertType.INFO);
    
    Logger.log(successMessage);
    
  } catch (error) {
    const errorMessage = `❌ 批量計算失敗: ${error.message}`;
    
    // 記錄錯誤
    logExecution('錯誤', '系統', '批量加班時數計算失敗', '控制中樞表', error.message);
    
    // 顯示錯誤給用戶
    SpreadsheetApp.getUi().alert('批量計算失敗', errorMessage, SpreadsheetApp.getUi().AlertType.ERROR);
    
    Logger.log(errorMessage);
  }
}

/**
 * 測試連線功能
 */
function testConnection() {
  const startTime = new Date();
  const CONTROL_CENTER_ID = '1fTQ3AZ93yP_q7oCncMASozScIJ36NlJwgc3vplr0nJI';
  
  try {
    // 記錄開始測試
    logExecution('測試', '系統', '開始測試批量加班計算功能', '測試模式', '讀取控制中樞表並測試第一位啟用員工');
    
    // 讀取控制中樞表的員工清單
    const activeEmployees = getActiveEmployeesFromControlCenter(CONTROL_CENTER_ID);
    
    if (activeEmployees.length === 0) {
      const testMessage = '⚠️ 測試失敗：沒有找到已啟用的員工';
      logExecution('測試', '系統', '測試失敗', '測試模式', '控制中樞表中無啟用員工');
      SpreadsheetApp.getUi().alert('測試結果', testMessage, SpreadsheetApp.getUi().AlertType.WARNING);
      return;
    }
    
    // 只測試第一位啟用的員工
    const testEmployee = activeEmployees[0];
    const overtimeRecords = calculateOvertimeForEmployee(testEmployee.fileId);
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    const testMessage = `✅ 測試成功！
━━━━━━━━━━━━━━━━━━━━
控制中樞表讀取: ✓
找到啟用員工: ${activeEmployees.length} 人
測試員工: ${testEmployee.name} (${testEmployee.id})
找到加班記錄: ${overtimeRecords.length} 筆
測試時間: ${duration.toFixed(2)} 秒

前3筆記錄範例:
${overtimeRecords.slice(0, 3).map((record, index) => 
  `${index + 1}. ${record.dateKey} - ${record.hours}小時 (${record.multiplier}倍)`
).join('\n')}

所有啟用員工清單:
${activeEmployees.map((emp, index) => 
  `${index + 1}. ${emp.name} (${emp.id})`
).join('\n')}`;
    
    // 記錄測試結果
    logExecution('測試', '系統', '批量測試完成', '測試模式', 
                `啟用員工: ${activeEmployees.length}人, 測試員工: ${testEmployee.name}, 找到記錄: ${overtimeRecords.length}筆, 耗時: ${duration.toFixed(2)}秒`);
    
    // 顯示測試結果給用戶
    SpreadsheetApp.getUi().alert('測試結果', testMessage, SpreadsheetApp.getUi().AlertType.INFO);
    
    Logger.log(testMessage);
    
  } catch (error) {
    const errorMessage = `❌ 測試失敗: ${error.message}`;
    
    // 記錄測試錯誤
    logExecution('錯誤', '系統', '批量測試失敗', '測試模式', error.message);
    
    // 顯示錯誤給用戶
    SpreadsheetApp.getUi().alert('測試失敗', errorMessage, SpreadsheetApp.getUi().AlertType.ERROR);
    
    Logger.log(errorMessage);
  }
}

/**
 * 顯示執行紀錄
 */
function showExecutionLog() {
  try {
    const executionSheet = getOrCreateExecutionLogSheet();
    const data = executionSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('執行紀錄', '目前沒有執行紀錄', SpreadsheetApp.getUi().AlertType.INFO);
      return;
    }
    
    // 顯示最近5筆記錄
    const recentRecords = data.slice(-6, -1).reverse(); // 排除標題行，取最後5筆，並反轉順序（最新在前）
    const logMessage = `最近執行紀錄：\n\n${recentRecords.map(row => 
      `${row[0]} [${row[1]}] ${row[2]}\n範圍: ${row[3]}\n詳情: ${row[4]}`
    ).join('\n\n')}`;
    
    SpreadsheetApp.getUi().alert('執行紀錄', logMessage, SpreadsheetApp.getUi().AlertType.INFO);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('錯誤', `無法讀取執行紀錄: ${error.message}`, SpreadsheetApp.getUi().AlertType.ERROR);
  }
}

/**
 * 計算單一員工的所有加班時數記錄
 */
function calculateOvertimeForEmployee(fileId) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    
    // 月份工作表名稱
    const monthlySheetNames = [
      '2025年1月工時紀錄', '2025年2月工時紀錄', '2025年3月工時紀錄',
      '2025年4月工時紀錄', '2025年5月工時紀錄', '2025年6月工時紀錄',
      '2025年7月', '2025年8月', '2025年9月', '2025年10月', '2025年11月'
    ];
    
    let allOvertimeRecords = [];
    
    for (const sheetName of monthlySheetNames) {
      try {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          console.log(`警告：找不到工作表 "${sheetName}"`);
          continue;
        }
        
        const workRecords = getWorkRecordsFromEmployeeSheet(sheet, sheetName);
        
        for (const record of workRecords) {
          const overtimeForDay = calculateOvertimeForDay(record);
          allOvertimeRecords.push(...overtimeForDay);
        }
      } catch (error) {
        const errorMsg = `處理工作表 "${sheetName}" 時發生錯誤: ${error.message}`;
        console.log(errorMsg);
        // 記錄到執行日誌以便用戶看到
        try {
          logExecution('錯誤', '系統', `工作表處理失敗: ${sheetName}`, '個人試算表', error.message);
        } catch (logError) {
          console.log(`日誌記錄失敗: ${logError.message}`);
        }
      }
    }
    
    return allOvertimeRecords;
  } catch (error) {
    throw new Error(`計算員工加班時數失敗: ${error.message}`);
  }
}

/**
 * 從日期字符串獲取星期
 */
function getDayOfWeekFromDate(dateString) {
  try {
    const date = new Date(dateString);
    const days = ['日', '一', '二', '三', '四', '五', '六'];
    return days[date.getDay()];
  } catch (error) {
    return '';
  }
}


// ===== 整合的 OvertimeCalculator 功能 =====

/**
 * 重命名函數以避免衝突
 */
function getWorkRecordsFromEmployeeSheet(sheet, sheetName) {
  const data = sheet.getDataRange().getValues();
  const records = [];
  
  // 判斷格式類型
  if (isMonthlyFormat1to5(sheetName)) {
    return parseMonthlyFormat1to5(data);
  } else if (isMonthlyFormat6(sheetName)) {
    return parseMonthlyFormat6(data);
  } else if (isMonthlyFormat7Plus(sheetName)) {
    return parseMonthlyFormat7Plus(data);
  }
  
  return records;
}

function isMonthlyFormat1to5(sheetName) {
  return sheetName.includes('年1月') || sheetName.includes('年2月') || 
         sheetName.includes('年3月') || sheetName.includes('年4月') || 
         sheetName.includes('年5月');
}

function isMonthlyFormat6(sheetName) {
  return sheetName.includes('年6月');
}

function isMonthlyFormat7Plus(sheetName) {
  return sheetName.includes('年7月') || sheetName.includes('年8月') || 
         sheetName.includes('年9月') || sheetName.includes('年10月') || 
         sheetName.includes('年11月') || sheetName.includes('年12月');
}

function parseMonthlyFormat1to5(data) {
  const records = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1] || !row[2]) continue;
    if (row[1].toString().includes('小計')) continue;
    
    const clockIn = new Date(row[1]);
    const clockOut = new Date(row[2]);
    const workHours = parseFloat(row[5]) || 0;
    
    if (workHours > 0 && !isNaN(clockIn.getTime()) && !isNaN(clockOut.getTime())) {
      if (clockOut < clockIn) {
        clockOut.setDate(clockOut.getDate() + 1);
      }
      
      const dateKey = formatDateKey(clockIn);
      
      records.push({
        date: clockIn.toDateString(),
        dateKey: dateKey,
        clockIn: clockIn,
        clockOut: clockOut,
        totalWorkHours: workHours,
        isWeekend: isWeekend(clockIn),
        isHoliday: false,
        note: row[6] || ''
      });
    }
  }
  
  return records;
}

function parseMonthlyFormat6(data) {
  const records = [];
  let currentDate = null;
  let currentSessions = [];
  
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    
    const cellB = row[1].toString();
    
    if ((cellB.includes('2025-') || cellB.includes('2025/')) && 
        (cellB.includes('例假日') || cellB.includes('休息日') || 
         (cellB.length > 10 && row[2]) ||  // 處理 "2025-06-30 15:30:00" 這種格式
         cellB.includes('日'))) {
      if (currentDate && currentSessions.length > 0) {
        const record = processDaySessionsFormat6(currentDate, currentSessions);
        if (record) records.push(record);
      }
      
      currentDate = {
        date: cellB,
        dayTotal: parseFloat(row[6]) || 0,
        note: row[7] || ''
      };
      currentSessions = [];
      
      // 6月格式：B欄和C欄直接是完整的時間戳
      if (cellB.includes(':') && row[2] && row[2].toString().includes(':')) {
        // 格式：B="2025-06-30 15:30:00", C="2025-06-30 19:30:59"
        currentSessions.push({
          clockIn: cellB,  // B欄就是上班時間
          clockOut: row[2], // C欄是下班時間
          hours: parseFloat(row[3]) || 0, // D欄是工作時數
          note: row[7] || ''
        });
      } else if (row[2] && row[3]) {
        // 原本的格式：C欄=上班, D欄=下班
        currentSessions.push({
          clockIn: row[2],
          clockOut: row[3],
          hours: parseFloat(row[4]) || 0,
          note: row[7] || ''
        });
      }
    } else if (currentDate) {
      // 處理同一日的其他時段
      const cellB2 = row[1] ? row[1].toString() : '';
      
      if (cellB2.includes(':') && row[2] && row[2].toString().includes(':')) {
        // 6月格式的其他時段：B="2025-06-30 13:30:00", C="2025-06-30 14:30:00"
        currentSessions.push({
          clockIn: cellB2,
          clockOut: row[2],
          hours: parseFloat(row[3]) || 0,
          note: row[7] || ''
        });
      } else if (row[2] && row[3]) {
        // 原本格式的其他時段
        currentSessions.push({
          clockIn: row[2],
          clockOut: row[3],
          hours: parseFloat(row[4]) || 0,
          note: row[7] || ''
        });
      }
    }
  }
  
  if (currentDate && currentSessions.length > 0) {
    const record = processDaySessionsFormat6(currentDate, currentSessions);
    if (record) records.push(record);
  }
  
  return records;
}

function parseMonthlyFormat7Plus(data) {
  const records = [];
  let currentDate = null;
  let currentSessions = [];
  
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    
    const cellB = row[1].toString();
    
    if (cellB.includes('2025')) {
      if (currentDate && currentSessions.length > 0) {
        const record = processDaySessionsFormat7Plus(currentDate, currentSessions);
        if (record) records.push(record);
      }
      
      currentDate = {
        date: cellB,
        dayType: row[2] || '',
        dayTotal: parseFloat(row[7]) || 0,
        overtimeHours: parseFloat(row[9]) || 0,
        note: row[10] || row[9] || ''
      };
      
      // 調試：記錄每日總工時的提取
      const debugMsg = `7+月格式提取: 日期=${cellB}, H欄dayTotal=${row[7]}, J欄overtime=${row[9]}`;
      console.log(debugMsg);
      logExecution('調試', '系統', debugMsg, '數據提取', '');
      currentSessions = [];
      
      // 檢查不同的欄位組合
      if (row[3] && row[4]) {
        // 標準7月格式：D欄=上班, E欄=下班
        currentSessions.push({
          clockIn: row[3],
          clockOut: row[4],
          hours: parseFloat(row[5]) || 0,
          workHours: parseFloat(row[7]) || 0,
          note: currentDate.note
        });
      } else if (row[2] && row[3]) {
        // 某些變體格式：C欄=上班, D欄=下班
        currentSessions.push({
          clockIn: row[2],
          clockOut: row[3],
          hours: parseFloat(row[4]) || 0,
          workHours: parseFloat(row[6]) || 0,
          note: currentDate.note
        });
      }
    } else if (currentDate && (row[3] && row[4])) {
      // 標準格式：D欄=上班, E欄=下班
      currentSessions.push({
        clockIn: row[3],
        clockOut: row[4],
        hours: parseFloat(row[5]) || 0,
        workHours: parseFloat(row[7]) || 0,
        note: row[10] || row[9] || ''
      });
    } else if (currentDate && (row[2] && row[3])) {
      // 變體格式：C欄=上班, D欄=下班
      currentSessions.push({
        clockIn: row[2],
        clockOut: row[3],
        hours: parseFloat(row[4]) || 0,
        workHours: parseFloat(row[6]) || 0,
        note: row[10] || row[9] || ''
      });
    }
  }
  
  if (currentDate && currentSessions.length > 0) {
    const record = processDaySessionsFormat7Plus(currentDate, currentSessions);
    if (record) records.push(record);
  }
  
  return records;
}

function processDaySessionsFormat6(dateInfo, sessions) {
  if (sessions.length === 0) return null;
  
  let date;
  const dateStr = dateInfo.date;
  if (dateStr.includes('例假日') || dateStr.includes('休息日')) {
    // 支援 2025-MM-DD 和 2025/MM/DD 格式
    const dateMatch = dateStr.match(/2025[-/]\d{2}[-/]\d{2}/);
    if (!dateMatch) return null;
    const normalizedDate = dateMatch[0].replace(/\//g, '-');
    date = new Date(normalizedDate);
  } else {
    const firstPart = dateStr.split(' ')[0];
    // 支援 2025-MM-DD 和 2025/MM/DD 格式
    const normalizedDate = firstPart.replace(/\//g, '-');
    date = new Date(normalizedDate);
  }
  
  let earliestClockIn = null;
  let latestClockOut = null;
  
  sessions.forEach(session => {
    if (session.note && session.note.includes('補休')) return;
    
    let clockIn, clockOut;
    
    // 檢查是否為完整時間戳格式 (如 "2025-06-30 15:30:00")
    if (session.clockIn.includes(' ') && session.clockIn.includes('-')) {
      clockIn = new Date(session.clockIn);
      clockOut = new Date(session.clockOut);
    } else {
      // 原本的格式：只有時間部分，需要加上日期
      clockIn = new Date(`${date.toDateString()} ${session.clockIn}`);
      clockOut = new Date(`${date.toDateString()} ${session.clockOut}`);
    }
    
    if (!earliestClockIn || clockIn < earliestClockIn) {
      earliestClockIn = clockIn;
    }
    if (!latestClockOut || clockOut > latestClockOut) {
      latestClockOut = clockOut;
    }
  });
  
  if (!earliestClockIn || !latestClockOut) return null;
  
  const dateKey = formatDateKey(date);
  
  return {
    date: date.toDateString(),
    dateKey: dateKey,
    clockIn: earliestClockIn,
    clockOut: latestClockOut,
    totalWorkHours: dateInfo.dayTotal,
    isWeekend: isWeekend(date) || dateStr.includes('休息日'),
    isHoliday: dateStr.includes('例假日'),
    note: dateInfo.note
  };
}

function processDaySessionsFormat7Plus(dateInfo, sessions) {
  if (sessions.length === 0) return null;
  
  const dateStr = dateInfo.date;
  let dateMatch = dateStr.match(/2025[.-]\d{2}[.-]\d{2}/);
  
  if (!dateMatch) {
    // 嘗試其他格式：2025.11.01 或 2025-10-01
    dateMatch = dateStr.match(/2025[.\-]\d{1,2}[.\-]\d{1,2}/);
  }
  
  if (!dateMatch) {
    console.log(`無法解析日期格式: ${dateStr}`);
    return null;
  }
  
  // 正則化日期格式
  let normalizedDate = dateMatch[0].replace(/\./g, '-');
  
  // 確保月份和日期為兩位數
  const dateParts = normalizedDate.split('-');
  if (dateParts.length === 3) {
    const year = dateParts[0];
    const month = dateParts[1].padStart(2, '0');
    const day = dateParts[2].padStart(2, '0');
    normalizedDate = `${year}-${month}-${day}`;
  }
  
  const date = new Date(normalizedDate);
  
  // 檢查日期是否有效
  if (isNaN(date.getTime())) {
    console.log(`無法創建有效日期: ${normalizedDate} 原始: ${dateStr}`);
    return null;
  }
  
  let earliestClockIn = null;
  let latestClockOut = null;
  let totalWorkHours = 0;
  
  sessions.forEach(session => {
    const noteStr = session.note ? session.note.toString() : '';
    if (noteStr && (noteStr.includes('補休') || noteStr.includes('午休') || noteStr.includes('晚休'))) {
      return;
    }
    
    let clockIn, clockOut;
    
    if (typeof session.clockIn === 'string') {
      if (session.clockIn.includes(':')) {
        const dateString = date.toISOString().split('T')[0]; // YYYY-MM-DD 格式
        clockIn = new Date(`${dateString} ${session.clockIn}`);
      } else {
        clockIn = new Date(session.clockIn);
      }
    } else {
      clockIn = new Date(session.clockIn);
    }
    
    if (typeof session.clockOut === 'string') {
      if (session.clockOut.includes(':')) {
        const dateString = date.toISOString().split('T')[0]; // YYYY-MM-DD 格式
        clockOut = new Date(`${dateString} ${session.clockOut}`);
      } else {
        clockOut = new Date(session.clockOut);
      }
    } else {
      clockOut = new Date(session.clockOut);
    }
    
    // 檢查時間是否有效
    if (isNaN(clockIn.getTime()) || isNaN(clockOut.getTime())) {
      console.log(`無法解析打卡時間: ${session.clockIn} - ${session.clockOut}`);
      return;
    }
    
    if (clockOut < clockIn) {
      clockOut.setDate(clockOut.getDate() + 1);
    }
    
    if (!earliestClockIn || clockIn < earliestClockIn) {
      earliestClockIn = clockIn;
    }
    if (!latestClockOut || clockOut > latestClockOut) {
      latestClockOut = clockOut;
    }
    
    totalWorkHours += session.workHours || session.hours || 0;
  });
  
  if (!earliestClockIn || !latestClockOut) return null;
  
  const finalWorkHours = dateInfo.dayTotal || totalWorkHours;
  const dateKey = formatDateKey(date);
  
  // 調試：記錄最終工時的決定過程
  const workHoursMsg = `${dateKey}: dayTotal=${dateInfo.dayTotal}, calculated=${totalWorkHours}, final=${finalWorkHours}`;
  console.log(workHoursMsg);
  logExecution('調試', '系統', workHoursMsg, '工時計算', '');
  
  const result = {
    date: date.toDateString(),
    dateKey: dateKey,
    clockIn: earliestClockIn,
    clockOut: latestClockOut,
    totalWorkHours: finalWorkHours,
    isWeekend: isWeekend(date) || dateInfo.dayType.includes('休息日'),
    isHoliday: dateInfo.dayType.includes('國定假日') || dateInfo.dayType.includes('特休'),
    note: dateInfo.note
  };
  
  const resultMsg = `成功創建記錄: ${dateKey}, 工時=${finalWorkHours}`;
  console.log(resultMsg);
  logExecution('調試', '系統', resultMsg, '記錄創建', '');
  
  return result;
}

function calculateOvertimeForDay(record) {
  const overtimeRecords = [];
  const workHours = record.totalWorkHours;
  
  // 記錄所有處理的記錄以便調試
  const logMsg = `處理日期 ${record.dateKey}: 工時=${workHours}小時`;
  console.log(logMsg);
  
  // 新規則：7.25小時以下不算加班
  if (workHours <= 7.25) {
    const skipMsg = `跳過 ${record.dateKey}: 工時${workHours}小時 ≤ 7.25小時 (緩衝時間)`;
    console.log(skipMsg);
    return [];
  }
  
  const clockInTime = new Date(record.clockIn);
  const regularEndTime = new Date(clockInTime.getTime() + 7.25 * 60 * 60 * 1000);
  
  if (record.isWeekend) {
    calculateWeekendOvertime(record, overtimeRecords);
  } else if (record.isHoliday) {
    calculateHolidayOvertime(record, overtimeRecords);
  } else {
    calculateWeekdayOvertime(record, overtimeRecords, regularEndTime);
  }
  
  return overtimeRecords;
}

function calculateWeekdayOvertime(record, overtimeRecords, regularEndTime) {
  const workHours = record.totalWorkHours;
  
  // 新規則：7-7.25小時為緩衝時間，不算加班
  if (workHours <= 7.25) {
    return; // 沒有加班
  }
  
  // 計算理論下班時間（上班時間 + 7.25小時，包含緩衝時間）
  const clockInTime = new Date(record.clockIn);
  const bufferEndTime = new Date(clockInTime.getTime() + 7.25 * 60 * 60 * 1000);
  let currentStartTime = new Date(bufferEndTime);
  
  if (workHours > 7.25 && workHours <= 8) {
    // 7.25~8小時：1倍
    const hours = workHours - 7.25;
    const endTime = new Date(currentStartTime.getTime() + hours * 60 * 60 * 1000);
    
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime),
      hours: parseFloat(hours.toFixed(2)),
      multiplier: 1,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  } else if (workHours > 8 && workHours <= 10) {
    // 7.25~8小時：1倍
    const hours1 = 8 - 7.25;  // 0.75小時
    const endTime1 = new Date(currentStartTime.getTime() + hours1 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime1),
      hours: parseFloat(hours1.toFixed(2)),
      multiplier: 1,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    // 8~10小時：1.34倍
    const hours2 = workHours - 8;
    const endTime2 = new Date(endTime1.getTime() + hours2 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime1),
      endDateTime: formatDateTime(endTime2),
      hours: parseFloat(hours2.toFixed(2)),
      multiplier: 1.34,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  } else if (workHours > 10) {
    // 7.25~8小時：1倍
    const hours1 = 8 - 7.25;  // 0.75小時
    const endTime1 = new Date(currentStartTime.getTime() + hours1 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime1),
      hours: parseFloat(hours1.toFixed(2)),
      multiplier: 1,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    // 8~10小時：1.34倍
    const endTime2 = new Date(endTime1.getTime() + 2 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime1),
      endDateTime: formatDateTime(endTime2),
      hours: 2,
      multiplier: 1.34,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    // 10小時以上：1.67倍
    const hours3 = workHours - 10;
    const endTime3 = new Date(endTime2.getTime() + hours3 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime2),
      endDateTime: formatDateTime(endTime3),
      hours: parseFloat(hours3.toFixed(2)),
      multiplier: 1.67,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  }
}

function calculateWeekendOvertime(record, overtimeRecords) {
  const workHours = record.totalWorkHours;
  let currentStartTime = new Date(record.clockIn);
  
  if (workHours > 0 && workHours <= 2) {
    // 第1-2小時：1.34倍（新規則）
    const endTime = new Date(currentStartTime.getTime() + workHours * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime),
      hours: parseFloat(workHours.toFixed(2)),
      multiplier: 1.34,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  } else if (workHours > 2 && workHours <= 8) {
    // 第1-2小時：1.34倍（新規則）
    const endTime1 = new Date(currentStartTime.getTime() + 2 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime1),
      hours: 2,
      multiplier: 1.34,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    const hours2 = workHours - 2;
    const endTime2 = new Date(endTime1.getTime() + hours2 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime1),
      endDateTime: formatDateTime(endTime2),
      hours: parseFloat(hours2.toFixed(2)),
      multiplier: 1.67,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  } else if (workHours > 8) {
    // 第1-2小時：1.34倍（新規則）
    const endTime1 = new Date(currentStartTime.getTime() + 2 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime1),
      hours: 2,
      multiplier: 1.34,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    const endTime2 = new Date(endTime1.getTime() + 6 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime1),
      endDateTime: formatDateTime(endTime2),
      hours: 6,
      multiplier: 1.67,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    const hours3 = Math.min(workHours - 8, 4);
    const endTime3 = new Date(endTime2.getTime() + hours3 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime2),
      endDateTime: formatDateTime(endTime3),
      hours: parseFloat(hours3.toFixed(2)),
      multiplier: 2.67,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  }
}

function calculateHolidayOvertime(record, overtimeRecords) {
  const workHours = record.totalWorkHours;
  let currentStartTime = new Date(record.clockIn);
  
  if (workHours > 0 && workHours <= 8) {
    const endTime = new Date(currentStartTime.getTime() + workHours * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime),
      hours: parseFloat(workHours.toFixed(2)),
      multiplier: 2,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  } else if (workHours > 8 && workHours <= 10) {
    const endTime1 = new Date(currentStartTime.getTime() + 8 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime1),
      hours: 8,
      multiplier: 2,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    const hours2 = workHours - 8;
    const endTime2 = new Date(endTime1.getTime() + hours2 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime1),
      endDateTime: formatDateTime(endTime2),
      hours: parseFloat(hours2.toFixed(2)),
      multiplier: 2.33,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  } else if (workHours > 10) {
    const endTime1 = new Date(currentStartTime.getTime() + 8 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(currentStartTime),
      endDateTime: formatDateTime(endTime1),
      hours: 8,
      multiplier: 2,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    const endTime2 = new Date(endTime1.getTime() + 2 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime1),
      endDateTime: formatDateTime(endTime2),
      hours: 2,
      multiplier: 2.33,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
    
    const hours3 = Math.min(workHours - 10, 2);
    const endTime3 = new Date(endTime2.getTime() + hours3 * 60 * 60 * 1000);
    overtimeRecords.push({
      dateKey: record.dateKey,
      startDateTime: formatDateTime(endTime2),
      endDateTime: formatDateTime(endTime3),
      hours: parseFloat(hours3.toFixed(2)),
      multiplier: 2.67,
      isWeekend: record.isWeekend,
      isHoliday: record.isHoliday
    });
  }
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

// ===== 輔助功能函數 =====

/**
 * 將加班記錄寫入「加班費紀錄總表」工作表
 */
function writeOvertimeRecordsToSheet(overtimeRecords) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let overtimeSheet = spreadsheet.getSheetByName('加班費紀錄總表');
    
    // 如果工作表不存在，創建它
    if (!overtimeSheet) {
      overtimeSheet = spreadsheet.insertSheet('加班費紀錄總表');
      // 設置標題行
      overtimeSheet.getRange(1, 1, 1, 5).setValues([[
        'ID', '開始時間', '結束時間', '時數', '倍率'
      ]]);
      // 格式化標題行
      const headerRange = overtimeSheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E8F4FD');
    }
    
    // 清空舊資料，保留標題行
    const lastRow = overtimeSheet.getLastRow();
    if (lastRow > 1) {
      overtimeSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    
    if (overtimeRecords.length === 0) {
      return 0;
    }
    
    // 生成序號並準備資料
    let dayCounter = {};
    const outputData = overtimeRecords.map(record => {
      // 生成 YYYYMMDD_序號 格式的ID
      const dateKey = record.dateKey;
      if (!dayCounter[dateKey]) {
        dayCounter[dateKey] = 1;
      } else {
        dayCounter[dateKey]++;
      }
      const recordId = `${dateKey}_${String(dayCounter[dateKey]).padStart(2, '0')}`;
      
      return [
        recordId,
        record.startDateTime,
        record.endDateTime,
        record.hours,
        record.multiplier
      ];
    });
    
    // 寫入資料
    overtimeSheet.getRange(2, 1, outputData.length, 5).setValues(outputData);
    
    // 設置數字格式
    overtimeSheet.getRange(2, 4, outputData.length, 2).setNumberFormat('0.00');
    
    return outputData.length;
  } catch (error) {
    throw new Error(`寫入加班記錄失敗: ${error.message}`);
  }
}

/**
 * 記錄執行日誌到「執行紀錄」工作表
 */
function logExecution(type, source, title, scope, details) {
  try {
    const executionSheet = getOrCreateExecutionLogSheet();
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 在第2行插入新記錄（保持最新記錄在最上面）
    executionSheet.insertRowAfter(1);
    executionSheet.getRange(2, 1, 1, 5).setValues([[
      timestamp,
      type || '系統',
      title || '執行',
      scope || '未指定',
      details || ''
    ]]);
    
    // 限制記錄行數（保留最近100筆）
    const maxRows = 101; // 1個標題行 + 100筆記錄
    const currentRows = executionSheet.getLastRow();
    if (currentRows > maxRows) {
      const rowsToDelete = currentRows - maxRows;
      executionSheet.deleteRows(maxRows + 1, rowsToDelete);
    }
    
  } catch (error) {
    console.log(`記錄執行日誌失敗: ${error.message}`);
  }
}

/**
 * 取得或創建執行紀錄工作表
 */
function getOrCreateExecutionLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let executionSheet = spreadsheet.getSheetByName('執行紀錄');
  
  if (!executionSheet) {
    executionSheet = spreadsheet.insertSheet('執行紀錄');
    
    // 設置標題行
    executionSheet.getRange(1, 1, 1, 5).setValues([[
      '時間', '屬性', '標題', '範圍', '詳細資訊'
    ]]);
    
    // 格式化標題行
    const headerRange = executionSheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#F0F8FF');
    
    // 設置欄寬
    executionSheet.setColumnWidth(1, 120); // 時間
    executionSheet.setColumnWidth(2, 80);  // 屬性
    executionSheet.setColumnWidth(3, 150); // 標題
    executionSheet.setColumnWidth(4, 100); // 範圍
    executionSheet.setColumnWidth(5, 250); // 詳細資訊
    
    // 凍結標題行
    executionSheet.setFrozenRows(1);
  }
  
  return executionSheet;
}

/**
 * 從控制中樞表讀取已啟用的員工清單
 */
function getActiveEmployeesFromControlCenter(controlCenterId) {
  try {
    const controlCenter = SpreadsheetApp.openById(controlCenterId);
    const employeeSheet = controlCenter.getSheetByName('員工清單');
    
    if (!employeeSheet) {
      throw new Error('找不到「員工清單」工作表');
    }
    
    const data = employeeSheet.getDataRange().getValues();
    const activeEmployees = [];
    
    // 從第2行開始（跳過標題行）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // 檢查是否有有效資料
      if (!row[0] || !row[1] || !row[2] || !row[4]) continue;
      
      const employeeId = row[0].toString().trim();
      const employeeName = row[1].toString().trim();
      const fileId = row[2].toString().trim();
      const status = row[4].toString().trim();
      
      // 只取狀態為「啟用」的員工
      if (status === '啟用') {
        activeEmployees.push({
          id: employeeId,
          name: employeeName,
          fileId: fileId,
          status: status
        });
      }
    }
    
    return activeEmployees;
  } catch (error) {
    throw new Error(`讀取控制中樞表員工清單失敗: ${error.message}`);
  }
}

/**
 * 將加班記錄寫入員工個人試算表的「加班費紀錄總表」
 */
function writeOvertimeRecordsToEmployeeSheet(overtimeRecords, employeeFileId) {
  try {
    const employeeSpreadsheet = SpreadsheetApp.openById(employeeFileId);
    let overtimeSheet = employeeSpreadsheet.getSheetByName('加班費紀錄總表');
    
    // 如果工作表不存在，創建它
    if (!overtimeSheet) {
      overtimeSheet = employeeSpreadsheet.insertSheet('加班費紀錄總表');
      // 設置標題行（與OvertimeCalculator.gs一致）
      overtimeSheet.getRange(1, 1, 1, 5).setValues([[
        'ID', '開始時間', '結束時間', '時數', '倍率'
      ]]);
      // 格式化標題行
      const headerRange = overtimeSheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E8F4FD');
    }
    
    // 清空舊資料，保留標題行
    const lastRow = overtimeSheet.getLastRow();
    if (lastRow > 1) {
      overtimeSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    
    if (overtimeRecords.length === 0) {
      return 0;
    }
    
    // 生成序號並準備資料
    let dayCounter = {};
    const outputData = overtimeRecords.map(record => {
      // 生成 YYYYMMDD_序號 格式的ID
      const dateKey = record.dateKey;
      if (!dayCounter[dateKey]) {
        dayCounter[dateKey] = 1;
      } else {
        dayCounter[dateKey]++;
      }
      const recordId = `${dateKey}_${String(dayCounter[dateKey]).padStart(2, '0')}`;
      
      return [
        recordId,
        record.startDateTime,
        record.endDateTime,
        record.hours,
        record.multiplier
      ];
    });
    
    // 寫入資料
    overtimeSheet.getRange(2, 1, outputData.length, 5).setValues(outputData);
    
    // 設置數字格式
    overtimeSheet.getRange(2, 4, outputData.length, 2).setNumberFormat('0.00');
    
    // 設置欄寬
    overtimeSheet.setColumnWidth(1, 120); // ID
    overtimeSheet.setColumnWidth(2, 150); // 開始時間
    overtimeSheet.setColumnWidth(3, 150); // 結束時間
    overtimeSheet.setColumnWidth(4, 80);  // 時數
    overtimeSheet.setColumnWidth(5, 80);  // 倍率
    
    return outputData.length;
  } catch (error) {
    throw new Error(`寫入員工試算表加班費紀錄失敗: ${error.message}`);
  }
}