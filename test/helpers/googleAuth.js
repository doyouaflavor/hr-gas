const { GoogleAuth } = require('google-auth-library');
const { google } = require('googleapis');

class GoogleSheetsAuth {
  constructor() {
    this.auth = null;
    this.sheets = null;
  }

  async initialize() {
    try {
      // 方法 1: 使用 Service Account JSON 檔案
      if (process.env.GOOGLE_SERVICE_ACCOUNT_PATH) {
        this.auth = new GoogleAuth({
          keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_PATH,
          scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
        });
      }
      // 方法 2: 使用環境變數中的認證資訊
      else if (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_PRIVATE_KEY) {
        const credentials = {
          client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
          private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
        };
        
        this.auth = new GoogleAuth({
          credentials,
          scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
        });
      }
      // 方法 3: 使用 API Key (僅適用於公開 Sheets)
      else if (process.env.GOOGLE_SHEETS_API_KEY) {
        this.sheets = google.sheets({ 
          version: 'v4', 
          auth: process.env.GOOGLE_SHEETS_API_KEY 
        });
        return this;
      }
      else {
        throw new Error('未設定 Google API 認證資訊。請檢查環境變數設定。');
      }

      // 建立 Sheets API 實例
      this.sheets = google.sheets({ version: 'v4', auth: this.auth });
      
      return this;
    } catch (error) {
      console.error('Google API 認證失敗:', error.message);
      throw error;
    }
  }

  async testConnection() {
    if (!this.sheets) {
      throw new Error('Google Sheets API 未初始化');
    }

    try {
      // 測試連接 - 嘗試取得 Sheet 的基本資訊
      const testSheetId = process.env.TEST_MASTER_SHEET_ID || process.env.TEST_SHEET_ID;
      const response = await this.sheets.spreadsheets.get({
        spreadsheetId: testSheetId
      });
      
      return {
        success: true,
        title: response.data.properties.title,
        sheets: response.data.sheets.map(sheet => sheet.properties.title)
      };
    } catch (error) {
      return {
        success: false,
        error: error.message
      };
    }
  }

  getClient() {
    return this.sheets;
  }
}

module.exports = GoogleSheetsAuth;