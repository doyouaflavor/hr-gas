const GoogleSheetsAuth = require('./googleAuth');

class SheetsReader {
  constructor() {
    this.auth = new GoogleSheetsAuth();
    this.sheets = null;
  }

  async initialize() {
    await this.auth.initialize();
    this.sheets = this.auth.getClient();
    return this;
  }

  /**
   * 讀取指定 Sheet 的所有工作表名稱
   */
  async getSheetNames(spreadsheetId) {
    try {
      const response = await this.sheets.spreadsheets.get({
        spreadsheetId
      });
      
      return response.data.sheets.map(sheet => ({
        id: sheet.properties.sheetId,
        title: sheet.properties.title,
        index: sheet.properties.index
      }));
    } catch (error) {
      throw new Error(`無法取得工作表清單: ${error.message}`);
    }
  }

  /**
   * 讀取指定範圍的資料
   */
  async getSheetData(spreadsheetId, range) {
    try {
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
        valueRenderOption: 'UNFORMATTED_VALUE',
        dateTimeRenderOption: 'FORMATTED_STRING'
      });
      
      return response.data.values || [];
    } catch (error) {
      throw new Error(`無法讀取 Sheet 資料 (${range}): ${error.message}`);
    }
  }

  /**
   * 讀取整個工作表的資料
   */
  async getFullSheetData(spreadsheetId, sheetName) {
    const range = `${sheetName}!A:Z`;
    return await this.getSheetData(spreadsheetId, range);
  }

  /**
   * 模擬 Apps Script 中的 scanMonthlySheets 功能
   */
  async scanMonthlySheets(spreadsheetId) {
    try {
      const allSheets = await this.getSheetNames(spreadsheetId);
      const monthlySheets = [];
      
      const monthPattern = /^(\d{1,2})月$/;
      
      for (const sheet of allSheets) {
        if (monthPattern.test(sheet.title)) {
          const monthMatch = sheet.title.match(monthPattern);
          monthlySheets.push({
            id: sheet.id,
            title: sheet.title,
            month: parseInt(monthMatch[1])
          });
        }
      }
      
      // 按月份排序
      monthlySheets.sort((a, b) => a.month - b.month);
      return monthlySheets;
    } catch (error) {
      throw new Error(`掃描月份工作表失敗: ${error.message}`);
    }
  }

  /**
   * 模擬 Apps Script 中的 extractOvertimeData 功能
   */
  async extractOvertimeData(spreadsheetId, sheetName, employeeId) {
    try {
      const data = await this.getFullSheetData(spreadsheetId, sheetName);
      const overtimeRecords = [];
      
      // 跳過標題行，從第二行開始處理
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // 確保行有足夠的欄位
        if (row.length < 12) continue;
        
        const date = row[1]; // B欄：日期
        const dayOfWeek = row[2]; // C欄：星期
        const dayType = row[3] || ''; // D欄：日期類型
        const overtimeHours = parseFloat(row[11]) || 0; // L欄：加班時數
        const note = row[12] || ''; // M欄：備註
        
        if (date && overtimeHours > 0) {
          const dateObj = this.parseDate(date);
          if (dateObj) {
            let overtimeType = this.determineOvertimeType(dayType, dayOfWeek);
            
            overtimeRecords.push({
              employeeId: employeeId,
              date: this.formatDate(dateObj),
              hours: overtimeHours,
              type: overtimeType,
              note: note,
              sourceMonth: sheetName,
              dayOfWeek: typeof dayOfWeek === 'number' ? this.getDayOfWeek(dateObj) : dayOfWeek
            });
          }
        }
      }
      
      return overtimeRecords;
    } catch (error) {
      throw new Error(`提取加班資料失敗 (${sheetName}): ${error.message}`);
    }
  }

  /**
   * 解析日期
   */
  parseDate(dateValue) {
    if (dateValue instanceof Date) {
      return dateValue;
    }
    
    if (typeof dateValue === 'string') {
      const parsed = new Date(dateValue);
      return isNaN(parsed.getTime()) ? null : parsed;
    }
    
    if (typeof dateValue === 'number') {
      // Google Sheets 數字日期格式 (自1900/1/1開始的天數)
      const baseDate = new Date(1899, 11, 30); // 1899/12/30
      return new Date(baseDate.getTime() + dateValue * 24 * 60 * 60 * 1000);
    }
    
    return null;
  }

  /**
   * 判斷加班類型
   */
  determineOvertimeType(dayType, dayOfWeek) {
    if (dayType.includes('例假日')) {
      return '例假日';
    } else if (dayType.includes('休息日') || dayOfWeek === '6' || dayOfWeek === '7') {
      return '假日';
    } else if (dayType.includes('上班日加班')) {
      return '上班日加班';
    } else if (dayType.includes('上班日')) {
      return '上班日';
    } else {
      return '平日';
    }
  }

  /**
   * 格式化日期
   */
  formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  /**
   * 取得星期幾
   */
  getDayOfWeek(date) {
    const days = ['日', '一', '二', '三', '四', '五', '六'];
    return days[date.getDay()];
  }

  /**
   * 測試連接
   */
  async testConnection() {
    return await this.auth.testConnection();
  }
}

module.exports = SheetsReader;