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