# 🔧 GAS 連線問題診斷指南

## 🚨 常見問題與解決方案

### 問題 1：網頁無法顯示資料（空白或「載入中...」）

**可能原因：**
- GAS 部署權限設定錯誤
- CORS 錯誤
- Google Sheet ID 未設定

**解決步驟：**

1. **檢查 GAS 部署設定**
   - 開啟 https://script.google.com
   - 找到你的專案
   - 點擊「部署」>「管理部署」
   - 點擊編輯圖示（鉛筆）
   - 確認設定：
     - ✅ 執行身分：**我** (你的 Google 帳號)
     - ✅ 誰有權存取：**任何人** (Anyone, even anonymous) ⚠️ 這個最重要！

2. **測試 GAS 網址**
   - 在瀏覽器直接開啟 GAS 網址（以 `/exec` 結尾）
   - 應該要看到 JSON 回應，例如：`{"status":"ok"}`
   - 如果看到 401/403 錯誤 → 權限設定錯誤
   - 如果看到 Google 登入頁面 → 需要設定為「任何人」

3. **檢查前端網址**
   - 開啟 `index.html` 或 `js/app.js`
   - 確認 `GAS_URL` 是正確的部署網址
   - 必須以 `/exec` 結尾，不是 `/dev`

---

### 問題 2：CORS 錯誤（"Failed to fetch"）

**症狀：**
```
NetworkError: Failed to fetch
Access to fetch at '...' from origin '...' has been blocked by CORS policy
```

**解決方案：**

1. **確認部署設定為「任何人」**（見問題 1）

2. **使用 POST 請求**（程式碼已處理）

3. **在瀏覽器測試**
   - 開啟 `test-gas.html` 進行診斷
   - 點擊「測試 POST 請求」按鈕

---

### 問題 3：Google Sheet ID 未設定

**症狀：**
- GAS 回應 `{"error":"找不到工作表：設備清單"}`

**解決方案：**

1. **取得 Sheet ID**
   - 開啟你的 Google Sheet
   - 從網址複製 ID：
   ```
   https://docs.google.com/spreadsheets/d/【這裡是 SPREADSHEET_ID】/edit
   ```

2. **設定 GAS**
   - 開啟 Google Apps Script 專案
   - 找到 `GOOGLE_APPS_SCRIPT.js`
   - 替換 `SPREADSHEET_ID` 變數：
   ```javascript
   const SPREADSHEET_ID = '你的實際 Sheet ID';
   ```

3. **確認工作表名稱**
   - 預設：`設備清單`
   - 如果不同，修改 `SHEET_NAME` 變數

---

### 問題 4：資料欄位對不上

**症狀：**
- 查詢得到資料但欄位錯誤
- 顯示 `undefined`

**解決方案：**

檢查你的 Google Sheet 欄位順序，調整 `GOOGLE_APPS_SCRIPT.js` 中的欄位索引：

```javascript
// 依照你的 Sheet 實際欄位順序調整
const colIndex = {
  'fix_no': 1,      // B 欄
  'fix_type': 2,    // C 欄
  'descrip': 3,     // D 欄
  'qty_asset': 4,   // E 欄
  'dept_id': 7,     // H 欄
  'keeper': 11,     // L 欄
  'status': 12,     // M 欄
  // ... 依此類推
};
```

---

## 🧪 快速測試流程

### 步驟 1：測試 GAS 連線
```bash
# 在瀏覽器開啟
open mt-equipment-system/test-gas.html
```

點擊「測試 POST 請求」按鈕：
- ✅ 看到 `{"status":"ok"}` → GAS 連線正常
- ❌ 看到錯誤 → 檢查部署設定

### 步驟 2：測試主網頁
```bash
# 用 Python 啟動本地伺服器
cd mt-equipment-system
python3 -m http.server 8080

# 開啟瀏覽器
open http://localhost:8080
```

### 步驟 3：檢查瀏覽器控制台
- 按 `F12` 開啟開發者工具
- 切換到 `Console` 分頁
- 查看是否有錯誤訊息

---

## 📋 完整部署檢查清單

- [ ] Google Sheet 已建立，包含「設備清單」工作表
- [ ] 欄位標題正確（fix_no, fix_type, descrip, keeper, status...）
- [ ] Google Apps Script 專案已建立
- [ ] `GOOGLE_APPS_SCRIPT.js` 程式碼已貼上
- [ ] `SPREADSHEET_ID` 已設定為正確的 Sheet ID
- [ ] 部署為 Web App
- [ ] 執行身分：我
- [ ] 誰有權存取：任何人
- [ ] 複製的網址以 `/exec` 結尾
- [ ] 前端 `GAS_URL` 已更新為正確網址
- [ ] 測試 `test-gas.html` 連線成功
- [ ] 測試主網頁可以載入資料

---

## 🆘 需要幫忙？

1. 開啟 `test-gas.html` 進行診斷
2. 檢查瀏覽器控制台（F12）的錯誤訊息
3. 確認 GAS 部署設定截圖
4. 提供錯誤訊息內容

---

🦐 由 蝦米東 提供支援
