require('dotenv').config();

// 設定測試超時時間
jest.setTimeout(30000);

// 全域測試配置
global.testConfig = {
  // 主控台試算表 ID（包含員工清單、加班記錄總表、補休申請總表）
  masterSheetId: process.env.TEST_MASTER_SHEET_ID,
  // 員工個人試算表 ID（包含各月份分頁）
  employeeSheetId: process.env.TEST_EMPLOYEE_SHEET_ID,
  // 向後相容的預設值
  sheetId: process.env.TEST_SHEET_ID || process.env.TEST_MASTER_SHEET_ID
};

// 在所有測試開始前的設定
beforeAll(() => {
  console.log('🧪 開始 Google Sheets 測試...');
});

// 在所有測試結束後的清理
afterAll(() => {
  console.log('✅ Google Sheets 測試完成');
});