# 🚀 HR 管理系統實作 - 下一步開發任務

## 📋 任務概述
基於已完成的 Code.js 實作，僅需要新增**個人補休表同步功能**和**例假日合規處理**。現有的員工加班資料同步、反向驗證、補休配對等核心功能已經完成。

## ✅ 已完成功能 (Code.js)
1. **員工加班資料同步** ✅ - `checkEmployeeData()`, `extractOvertimeData()`
2. **反向驗證** ✅ - `validateOvertimeRecords()`  
3. **補休配對機制** ✅ - `matchLeaveWithOvertime()`, `allocateLeaveToOvertime()`
4. **主控台觸發** ✅ - `checkAllEmployees()`
5. **工具函數** ✅ - 完整的輔助函數庫

## 🎯 待實作功能 (新增需求)
1. **個人補休表同步**：從員工個人補休表同步到中樞表
2. **例假日合規處理**：例假日加班標記為不可補休
3. **補休編號回寫**：將配對結果回寫到員工個人表格

## 📝 TDD 實作步驟

### **階段 1: 更新現有功能支援例假日處理**

#### Step 1.1: 更新例假日加班處理邏輯
**目標：** 修改 `addOvertimeRecord()` 和 `extractOvertimeData()` 支援例假日不可補休

**測試檔案：** `test/holiday-overtime.test.js`

**Red Phase (寫失敗測試):**
- [ ] `test('例假日加班應標記為不可補休狀態')`
- [ ] `test('例假日加班的剩餘可補休應為0')`
- [ ] `test('例假日加班狀態應為例假日-僅發加班費')`
- [ ] `test('補休配對應排除例假日加班記錄')`

**Green Phase (修改現有函數):**
```javascript
// 修改 Code.js 中的 addOvertimeRecord() 函數
// 增加例假日判斷邏輯，設定剩餘可補休=0，狀態=例假日-僅發加班費
```

**修改重點：**
- [ ] 在 `addOvertimeRecord()` 中加入例假日判斷
- [ ] 在 `allocateLeaveToOvertime()` 中排除例假日記錄
- [ ] 更新狀態管理邏輯

### **階段 2: 新增個人補休表同步功能**

#### Step 2.1: 個人補休表同步功能 (TDD)
**目標：** 新增 `syncPersonalLeaveRequests()` 函數

**測試檔案：** `test/personal-leave-sync.test.js`

**Red Phase (寫失敗測試):**
- [ ] `test('應該能掃描個人補休表分頁')`
- [ ] `test('應該能產生補休編號 LV-YYYYMMDD-E001-1')`
- [ ] `test('應該能自動配對可用的加班記錄')`
- [ ] `test('應該拒絕例假日加班記錄配對')`
- [ ] `test('應該處理補休時數超過可用加班時數的錯誤')`
- [ ] `test('應該在中樞表建立補休記錄並加入核選方框')`
- [ ] `test('應該回寫補休編號到個人補休表')`

**Green Phase (實作新函數):**
```javascript
// 在 Code.js 中新增以下函數：
function syncPersonalLeaveRequests() { }
function generateLeaveId(date, employeeId) { }
function scanPersonalLeaveSheet(fileId) { }
function addLeaveRequestToMaster(leaveRequest) { }
function writeLeaveIdToPersonalSheet(fileId, rowIndex, leaveId) { }
```

**實作重點：**
- [ ] 掃描所有員工的「個人補休表」分頁
- [ ] 產生補休編號 `LV-YYYYMMDD-員工編號-流水號`
- [ ] 自動配對可用的加班記錄（排除例假日）
- [ ] 同步到主控台的「補休申請總表」
- [ ] 資料驗證與錯誤處理
- [ ] 回寫補休編號到個人表格

#### Step 2.2: 補休編號回寫功能 (TDD)
**目標：** 修改現有的補休配對流程，加入回寫機制

**測試檔案：** `test/leave-id-writeback.test.js`

**Red Phase (寫失敗測試):**
- [ ] `test('配對完成後應回寫用掉補休編號到加班記錄總表')`
- [ ] `test('應該同步補休編號到員工個人月份分頁')`
- [ ] `test('多筆補休使用同一加班記錄時應用逗號分隔')`

**Green Phase (修改現有函數):**
```javascript
// 修改 Code.js 中的以下函數：
// 1. updateLeaveRecord() - 增加回寫到加班記錄總表
// 2. 新增 writeLeaveIdToEmployeeSheet() 函數
// 3. 修改 allocateLeaveToOvertime() 支援回寫
```

**修改重點：**
- [ ] 在 `allocateLeaveToOvertime()` 中加入補休編號回寫
- [ ] 更新「加班記錄總表」的「用掉補休編號」欄位
- [ ] 同步回寫到員工個人月份分頁
- [ ] 處理多對一的補休-加班關係

### **階段 3: 整合新功能到主流程**

#### Step 3.1: 更新主控台觸發流程
**目標：** 在 `checkAllEmployees()` 中加入新的同步步驟

**修改重點：**
- [ ] 在 Step 2 之後加入 `syncPersonalLeaveRequests()`
- [ ] 確保執行順序正確：加班同步 → 個人補休同步 → 配對處理
- [ ] 更新執行報告包含新功能的統計

**執行順序：**
```
Step 1: 員工加班資料同步 (已完成)
Step 2: 反向驗證 (已完成)  
Step 3: 個人補休表同步 (新增)
Step 4: 補休與加班配對 (已完成，需修改支援回寫)
```

### **階段 4: 測試與驗證**

#### Step 4.1: 建立測試資料
- [ ] 在測試用員工試算表中建立「個人補休表」分頁
- [ ] 準備測試資料：包含例假日加班、一般加班、補休申請
- [ ] 設定測試案例涵蓋所有邊界條件

#### Step 4.2: 整合測試
- [ ] 測試完整流程：從個人表格到主控台的雙向同步
- [ ] 驗證例假日加班不參與補休配對
- [ ] 確認補休編號正確回寫到所有相關位置
- [ ] 測試錯誤處理和邊界情況

#### Step 4.3: 效能優化
- [ ] 批次讀寫操作優化
- [ ] 減少 API 呼叫次數
- [ ] 確保大量資料處理效能

## 🔧 實作細節

### **關鍵修改點**

#### 1. 修改 `addOvertimeRecord()` (Code.js:182)
```javascript
// 現有邏輯需要加入例假日判斷
const newRow = [
  overtimeId,
  record.employeeId, 
  employeeName,
  record.date,
  record.dayOfWeek,
  record.type,
  record.hours,
  0, // 已使用補休
  record.type === '例假日' ? 0 : record.hours, // 剩餘可補休 - 例假日為0
  record.type === '例假日' ? '例假日-僅發加班費' : '未使用', // 狀態
  record.sourceMonth,
  '', // 用掉補休編號
  ''  // 錯誤提示
];
```

#### 2. 修改 `allocateLeaveToOvertime()` (Code.js:294)
```javascript
// 在第294行附近加入例假日過濾
for (let i = 1; i < data.length; i++) {
  const row = data[i];
  if (row[1] === employeeId && 
      parseFloat(row[8]) > 0 && 
      row[9] !== '例假日-僅發加班費') { // 排除例假日
    availableRecords.push({...});
  }
}
```

#### 3. 新增函數架構
```javascript
// 在 Code.js 末尾新增以下函數：

function syncPersonalLeaveRequests() {
  // 掃描所有員工的個人補休表並同步到中樞表
}

function generateLeaveId(date, employeeId) {
  // 產生補休編號 LV-YYYYMMDD-E001-1
}

function scanPersonalLeaveSheet(fileId) {
  // 掃描單一員工的個人補休表分頁
}

function addLeaveRequestToMaster(leaveRequest) {
  // 將補休申請加入主控台補休申請總表
}

function writeLeaveIdToPersonalSheet(fileId, rowIndex, leaveId) {
  // 回寫補休編號到個人補休表
}

function writeLeaveIdToEmployeeSheet(employeeId, date, leaveId) {
  // 回寫補休編號到員工個人月份分頁
}
```

## 📊 完成指標

### **必須完成的功能**
- [ ] 例假日加班標記為不可補休 (修改現有函數)
- [ ] 個人補休表同步功能 (新增 `syncPersonalLeaveRequests()`)
- [ ] 補休編號回寫機制 (修改現有配對流程)
- [ ] 更新主流程整合新功能

### **測試完成度**
- [ ] 例假日處理測試通過
- [ ] 個人補休表同步測試通過  
- [ ] 補休編號回寫測試通過
- [ ] 整合流程測試通過

## 🎯 立即行動

### **第一步：建立測試環境**
```bash
# 在員工測試試算表中新增「個人補休表」分頁
# 準備測試資料包含例假日加班記錄
```

### **第二步：TDD 開發**
1. 先寫例假日處理的失敗測試
2. 修改現有函數支援例假日判斷
3. 寫個人補休表同步的失敗測試  
4. 實作新的同步函數
5. 寫補休編號回寫的失敗測試
6. 修改配對流程支援回寫

### **第三步：整合驗證**
1. 更新 `checkAllEmployees()` 主流程
2. 執行完整的整合測試
3. 在實際 Google Sheets 環境驗證

---

**重點：** 基於現有的 Code.js 實作，只需要**新增 3 個新函數**和**修改 2 個現有函數**，就能完成所有新需求！ 🚀