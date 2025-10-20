const SheetsReader = require('./helpers/sheetsReader');

describe('Google Sheets 連接測試', () => {
  let reader;

  beforeAll(async () => {
    reader = new SheetsReader();
    await reader.initialize();
  });

  test('應該能成功連接到 Google Sheets API', async () => {
    const result = await reader.testConnection();
    
    expect(result.success).toBe(true);
    expect(result.title).toBeDefined();
    expect(result.sheets).toBeInstanceOf(Array);
    
    console.log('📊 Sheet 標題:', result.title);
    console.log('📄 可用工作表:', result.sheets);
  });

  test('應該能取得所有工作表名稱', async () => {
    const sheets = await reader.getSheetNames(global.testConfig.sheetId);
    
    expect(sheets).toBeInstanceOf(Array);
    expect(sheets.length).toBeGreaterThan(0);
    
    sheets.forEach(sheet => {
      expect(sheet).toHaveProperty('id');
      expect(sheet).toHaveProperty('title');
      expect(sheet).toHaveProperty('index');
    });
    
    console.log('📋 工作表清單:', sheets.map(s => s.title));
  });
});