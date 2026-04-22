# 歷史紀錄修復指南

## 問題診斷步驟

### 1️⃣ 檢查 Google Sheet 是否有「歷史紀錄」工作表

開啟你的 Google Sheet：
https://docs.google.com/spreadsheets/d/1zW8SfCm8YtKwSfEnxqACn78TJaY4XIY5YL-OPZHliGY/edit

**檢查：**
- 底部工作表分頁是否有「歷史紀錄」？
- 如果沒有，請手動新增一個工作表，命名為 `歷史紀錄`
- 第一列應該有標題：`時間戳 | 動作 | 設備編號 | 設備名稱 | 借用人 | 保管人 | 借用日期 | 預計歸還 | 實際歸還/確認日期`

### 2️⃣ 更新 Google Apps Script 程式碼

1. 在 Google Sheet 中：**擴展功能** → **Apps Script**
2. 複製以下完整程式碼並覆蓋現有內容：
   - 檔案位置：`GOOGLE_APPS_SCRIPT.js`（從 GitHub 複製最新版本）
   - GitHub: https://github.com/lobbster1234-ai/mt-equipment-system/blob/main/GOOGLE_APPS_SCRIPT.js

3. **重新部署**：
   - 點擊 **部署** → **管理部署**
   - 點擊鉛筆圖示（編輯）
   - 版本：**新增版本**
   - 執行身分：**我** (stella_fan)
   - 誰有權存取：**任何人**
   - 點擊 **部署**

4. 複製新的部署網址（以 `/exec` 結尾）

### 3️⃣ 更新前端 GAS_URL

如果 GAS 部署網址有變更，需要更新：

**檔案：** `js/app.js`
```javascript
const GAS_URL = 'https://script.google.com/macros/s/YOUR_NEW_DEPLOYMENT_ID/exec';
```

### 4️⃣ 測試歷史紀錄功能

1. 開啟網頁：https://lobbster1234-ai.github.io/mt-equipment-system/
2. 按 **F12** 開啟開發者工具
3. 切換到 **Console** 分頁
4. 點擊「📜 歷史紀錄」分頁
5. 觀察 Console 輸出：
   - 應該看到 `歷史紀錄請求網址：...`
   - 應該看到 `歷史紀錄回應：...`
   - 如果看到錯誤訊息，請回報

### 5️⃣ 常見問題

#### ❌ 問題：Console 顯示 401/403 錯誤
**原因：** GAS 部署權限設定錯誤
**解法：** 確認部署設定為「誰有權存取：任何人」

#### ❌ 問題：顯示「目前沒有歷史紀錄」
**原因：** 歷史紀錄工作表是空的（正常，如果還沒借用/歸還過）
**解法：** 先進行一次借用操作，然後再查看歷史紀錄

#### ❌ 問題：CORS 錯誤 / Failed to fetch
**原因：** GAS 網址錯誤或部署設定問題
**解法：** 
- 確認 GAS_URL 正確（以 `/exec` 結尾，不是 `/dev`）
- 在瀏覽器直接開啟 GAS_URL，應該看到 JSON 回應

#### ❌ 問題：工作表不存在
**原因：** 「歷史紀錄」工作表尚未建立
**解法：** 
- 手動在 Google Sheet 新增工作表，命名為 `歷史紀錄`
- 或進行一次借用操作，系統會自動建立

## 修正內容摘要

### 後端修正（GOOGLE_APPS_SCRIPT.js）
- ✅ `borrowEquipment` 加入 `logHistory('borrow', ...)`
- ✅ `returnEquipment` 加入 `logHistory('return', ...)`
- ✅ `confirmReturn` 加入 `logHistory('confirm', ...)`
- ✅ `queryHistory` 修正日期欄位對應邏輯
- ✅ `queryHistory` 修正參數名稱（`actionType` 避免衝突）

### 前端修正（js/app.js）
- ✅ `searchHistory` 使用正確的參數名稱 `actionType`
- ✅ `renderHistory` 支援三種動作類型顯示（借用/歸還/確認）

## 需要協助？

如果以上步驟都完成但歷史紀錄仍然無法顯示，請提供：
1. F12 Console 的完整錯誤訊息
2. Google Sheet 工作表截圖
3. GAS 部署網址（確認格式正確）
