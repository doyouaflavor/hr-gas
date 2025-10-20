require('dotenv').config();

// 設定測試超時時間
jest.setTimeout(30000);

// 全域測試配置
global.testConfig = {
  sheetId: process.env.TEST_SHEET_ID || '1fTQ3AZ93yP_q7oCncMASozScIJ36NlJwgc3vplr0nJI'
};

// 在所有測試開始前的設定
beforeAll(() => {
  console.log('🧪 開始 Google Sheets 測試...');
});

// 在所有測試結束後的清理
afterAll(() => {
  console.log('✅ Google Sheets 測試完成');
});