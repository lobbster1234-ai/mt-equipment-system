# 📧 電子郵件通知設定指南

## ✅ 已完成的功能

### 借用流程
1. **借用者**點擊「借用」按鈕
2. 填寫姓名和預計歸還日期
3. 送出後設備狀態改為「借用中」
4. 📧 **自動寄信給 Keeper**通知有人借用

### 歸還流程
1. **借用者**點擊「歸還」按鈕
2. 確認歸還日期
3. 送出後系統記錄歸還日期
4. 📧 **自動寄信給 Keeper**通知歸還
5. **Keeper**到網頁點擊「✅ 確認歸還」
6. 設備狀態改為「可借用」
7. 📧 **自動寄信給 Keeper**確認完成

---

## 🔧 設定步驟

### 步驟 1：更新 GAS 程式碼

1. **開啟 Google Apps Script**
2. **完全替換程式碼**為 `GOOGLE_APPS_SCRIPT.js` 的內容
3. **設定 Sheet ID**：
   ```javascript
   const SPREADSHEET_ID = '1zW8SfCm8YtKwSfEnxqACn78TJaY4XIY5YL-OPZHliGY';
   ```
4. **儲存**

### 步驟 2：設定 Keeper 電子郵件對照表

在 `GOOGLE_APPS_SCRIPT.js` 中找到 `getKeeperEmail` 函式：

```javascript
function getKeeperEmail(keeperName) {
  // 方法 1：如果 keeper 欄位直接就是 email
  if (keeperName && keeperName.includes('@')) {
    return keeperName;
  }
  
  // 方法 2：建立 keeper 對照表
  const keeperEmailMap = {
    // 在這裡添加你的 keeper 電子郵件
    '王小明': 'wang@example.com',
    '李小華': 'li@example.com',
    '張小美': 'chang@example.com',
  };
  
  return keeperEmailMap[keeperName] || null;
}
```

**兩種方式：**

#### 方式 A：在 Sheet 中直接填寫 Email（推薦）
- 在 E 欄（keeper）直接填寫電子郵件地址
- 例如：`wang@example.com`
- 系統會自動辨識並寄信

#### 方式 B：使用對照表
- 在 keeper 欄位填寫中文名稱
- 在 `keeperEmailMap` 中建立名稱到 email 的對照
- 適合 keeper 名稱固定且人數不多的情況

### 步驟 3：重新部署

1. 點擊「部署」>「管理部署」
2. 點擊編輯圖示（鉛筆）
3. 版本選擇「**新增版本**」
4. 點擊「部署」
5. 確認網址是以 `/exec` 結尾

### 步驟 4：測試

1. **重新整理網頁**（F5）
2. 點擊「借用」按鈕
3. 填寫資料並送出
4. 檢查 Keeper 的 email 是否收到通知

---

## 📋 Google Sheet 欄位說明

| 欄 | 欄位名 | 說明 | 範例 |
|----|--------|------|------|
| A | fix_type | 設備類型 | 儀器設備 |
| B | fix_no | 設備編號 | MT-2024-001 |
| C | device_name | 設備名稱 | 示波器 Tektronix |
| D | qty_asset | 數量 | 1 |
| E | keeper | 保管人 | `wang@example.com` 或 王小明 |
| F | status | 狀態 | available/borrowed |
| G | borrower | 借用人 | 李小華 |
| H | dt_borrow | 借用日期 | 2024-01-15 |
| I | dt_due | 預計歸還 | 2024-01-22 |
| J | dt_return | 歸還日期 | 2024-01-20 |
| K | return_confirmed | 已確認歸還 | TRUE/FALSE |

---

## ⚠️ 重要提醒

### 電子郵件權限
- GAS 需要使用 `MailApp.sendEmail()` 發送郵件
- 第一次執行時會要求郵件權限
- 點擊「審查權限」>「進階」>「允許」

### 寄信限制
- Google 帳號每天有寄信上限（約 100-500 封/天）
- 一般使用不會超過限制
- 如果超過限制，會暫時無法寄信

### 測試模式
如果想關閉寄信功能測試，可以修改：
```javascript
const EMAIL_CONFIG = {
  enabled: false,  // 設為 false 關閉寄信
  // ...
};
```

---

## 🎨 網頁按鈕說明

| 按鈕 | 顏色 | 功能 |
|------|------|------|
| 借用 | 藍色 | 借用設備，填寫借用人和預計歸還日期 |
| 歸還 | 綠色 | 歸還設備，通知 Keeper |
| 確認歸還 | 青色 | Keeper 確認收到設備，完成歸還流程 |

---

## 📧 郵件範本

### 借用通知
```
親愛的 [Keeper] 您好：

有人借用了您保管的設備，詳情如下：

📦 設備編號：[fix_no]
📝 設備名稱：[device_name]
👤 借用人：[borrower]
📅 借用日期：[dt_borrow]
⏰ 預計歸還：[dt_due]

請留意設備歸還狀況。
```

### 歸還通知
```
親愛的 [Keeper] 您好：

您保管的設備已被歸還，請確認收到：

📦 設備編號：[fix_no]
📝 設備名稱：[device_name]
👤 原借用人：[borrower]
📅 歸還日期：[dt_return]

⚠️ 請前往以下網頁確認歸還：
[網頁網址]
```

### 歸還確認
```
親愛的 [Keeper] 您好：

您已確認收到歸還的設備：

📦 設備編號：[fix_no]
📝 設備名稱：[device_name]

✅ 設備狀態已更新為「可借用」

感謝您的配合！
```

---

## 🐛 問題排除

### Q: 沒有收到郵件？
A: 
1. 檢查 keeper 欄位是否有正確的 email
2. 檢查垃圾郵件夾
3. 確認 GAS 有郵件權限
4. 查看 GAS 的「執行記錄」是否有錯誤

### Q: 郵件寄給錯誤的人？
A: 檢查 `keeperEmailMap` 對照表或 keeper 欄位的內容

### Q: 按鈕沒有出現？
A: 
1. 重新整理網頁（F5）
2. 檢查瀏覽器控制台（F12）是否有錯誤
3. 確認 GAS 程式碼已更新並重新部署

---

🦐 由 蝦米東 提供支援
