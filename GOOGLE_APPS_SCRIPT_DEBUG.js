// =============================================
// MT 設備系統 - Google Apps Script 除錯版本
// =============================================
// 用途：診斷工作表名稱問題
// 使用方式：
// 1. 在 GAS 中完全替換原有程式碼
// 2. 替換 SPREADSHEET_ID
// 3. 儲存並重新部署
// 4. 直接在瀏覽器開啟 GAS 網址
// 5. 查看回傳的 JSON，確認工作表名稱
// =============================================

// ⚠️⚠️⚠️ 請替換成你的實際 Sheet ID ⚠️⚠️⚠️
const SPREADSHEET_ID = '請替換成你的 Google Sheet ID';

/**
 * GET 請求處理 - 除錯模式
 */
function doGet(e) {
  try {
    // 開啟 Sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 取得所有工作表
    const sheets = ss.getSheets();
    
    // 建立工作表資訊列表
    const sheetInfo = sheets.map((s, index) => ({
      index: index,
      name: s.getName(),
      rowCount: s.getLastRow(),
      columnCount: s.getLastColumn()
    }));
    
    // 回傳除錯資訊
    const result = {
      status: 'success',
      message: '✅ GAS 連線成功！請檢查下面的工作表列表',
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetName: ss.getName(),
      totalSheets: sheets.length,
      availableSheets: sheetInfo,
      instruction: '請比對 availableSheets 中的 name 欄位，確認你的工作表名稱是否完全一致（包含空格和大小寫）'
    };
    
    return ContentService.createTextOutput(JSON.stringify(result, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    const errorResult = {
      status: 'error',
      error: err.message,
      spreadsheetId: SPREADSHEET_ID,
      possibleIssues: [
        'Sheet ID 可能錯誤或不存在',
        'GAS 沒有權限讀取這個 Sheet',
        'Sheet 可能已被刪除'
      ]
    };
    
    return ContentService.createTextOutput(JSON.stringify(errorResult, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * POST 請求處理 - 除錯模式（同 GET）
 */
function doPost(e) {
  return doGet(e);
}
