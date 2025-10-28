# Google Sheets 測試框架

這個測試框架提供完整的 Node.js 環境來測試 Google Sheets 整合功能，無需在 Apps Script 環境中執行。

## 🚀 快速開始

### 1. 安裝依賴
```bash
npm install
```

### 2. 設定環境變數
複製 `.env.example` 為 `.env` 並填入您的認證資訊：

```bash
cp .env.example .env
```

編輯 `.env` 檔案，設定以下其中一種認證方式：

#### 方法一：Service Account（推薦）
```env
GOOGLE_SERVICE_ACCOUNT_EMAIL=your_service_account@your_project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\nyour_private_key_here\n-----END PRIVATE KEY-----"
# 主控台試算表 ID（包含員工清單、加班記錄總表、補休申請總表）
TEST_MASTER_SHEET_ID=your_master_sheet_id_here
# 員工個人試算表 ID（包含各月份分頁）
TEST_EMPLOYEE_SHEET_ID=your_employee_sheet_id_here
```

#### 方法二：Service Account JSON 檔案
```env
GOOGLE_SERVICE_ACCOUNT_PATH=/path/to/your/service-account.json
# 主控台試算表 ID（包含員工清單、加班記錄總表、補休申請總表）
TEST_MASTER_SHEET_ID=your_master_sheet_id_here
# 員工個人試算表 ID（包含各月份分頁）
TEST_EMPLOYEE_SHEET_ID=your_employee_sheet_id_here
```

#### 方法三：API Key（僅適用於公開 Sheets）
```env
GOOGLE_SHEETS_API_KEY=your_api_key_here
# 主控台試算表 ID（包含員工清單、加班記錄總表、補休申請總表）
TEST_MASTER_SHEET_ID=your_master_sheet_id_here
# 員工個人試算表 ID（包含各月份分頁）
TEST_EMPLOYEE_SHEET_ID=your_employee_sheet_id_here
```

### 3. 執行測試
```bash
# 執行所有測試
npm test

# 監控模式
npm run test:watch

# 產生覆蓋率報告
npm run test:coverage
```

## 📁 檔案結構

```
test/
├── README.md                 # 說明文件
├── setup.js                  # 測試環境設定
├── helpers/
│   ├── googleAuth.js         # Google API 認證模組
│   └── sheetsReader.js       # Sheets 資料讀取模組
├── connection.test.js        # 連接測試
├── dataExtraction.test.js    # 資料提取測試
└── integration.test.js       # 整合測試
```

## 🧪 測試類型

### 1. 連接測試 (`connection.test.js`)
- 測試 Google Sheets API 連接
- 驗證認證設定
- 取得工作表清單

### 2. 資料提取測試 (`dataExtraction.test.js`)
- 掃描月份工作表
- 讀取指定範圍資料
- 提取加班記錄
- 讀取完整工作表

### 3. 整合測試 (`integration.test.js`)
- 模擬完整的 `checkEmployeeData` 流程
- 資料驗證流程
- 效能測試
- 統計分析

## 🔧 核心模組

### GoogleSheetsAuth
負責處理 Google API 認證，支援多種認證方式。

### SheetsReader
提供與 Apps Script 相容的資料讀取功能：
- `scanMonthlySheets()` - 掃描月份工作表
- `extractOvertimeData()` - 提取加班資料
- `getSheetData()` - 讀取指定範圍
- `getFullSheetData()` - 讀取完整工作表

## 📊 測試範例

### 基本用法
```javascript
const SheetsReader = require('./helpers/sheetsReader');

const reader = new SheetsReader();
await reader.initialize();

// 掃描月份工作表
const monthlySheets = await reader.scanMonthlySheets(sheetId);

// 提取加班資料
const overtimeData = await reader.extractOvertimeData(
  sheetId, 
  '1月', 
  'EMP001'
);
```

### 進階用法
```javascript
// 讀取指定範圍
const data = await reader.getSheetData(sheetId, 'A1:Z100');

// 讀取完整工作表
const fullData = await reader.getFullSheetData(sheetId, '1月');

// 測試連接
const result = await reader.testConnection();
```

## 🔐 認證設定

### Service Account 設定步驟

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立或選擇專案
3. 啟用 Google Sheets API
4. 建立 Service Account
5. 下載 JSON 金鑰檔案或複製認證資訊
6. 將 Service Account 電子信箱加入 Google Sheets 共用清單

### API Key 設定步驟

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 啟用 Google Sheets API
3. 建立 API 金鑰
4. 設定 Sheets 為公開可讀取

## 🚨 故障排除

### 常見錯誤

1. **403 Forbidden**: 檢查 Service Account 是否有 Sheets 存取權限
2. **404 Not Found**: 確認 Sheet ID 正確
3. **認證失敗**: 檢查環境變數設定

### 除錯模式
```bash
# 顯示詳細錯誤訊息
npm test -- --verbose
```

## 🔄 CI/CD 整合

在 CI/CD 環境中使用時，建議：

1. 將認證資訊設為環境變數
2. 使用 Service Account 認證
3. 設定適當的測試超時時間
4. 產生測試報告

```yaml
# GitHub Actions 範例
env:
  GOOGLE_SERVICE_ACCOUNT_EMAIL: ${{ secrets.GOOGLE_SERVICE_ACCOUNT_EMAIL }}
  GOOGLE_PRIVATE_KEY: ${{ secrets.GOOGLE_PRIVATE_KEY }}
  TEST_MASTER_SHEET_ID: ${{ secrets.TEST_MASTER_SHEET_ID }}
  TEST_EMPLOYEE_SHEET_ID: ${{ secrets.TEST_EMPLOYEE_SHEET_ID }}
```