// =============================================
// MT 設備系統 - Google Apps Script 後端（完整版）
// =============================================
// 功能：查詢、登記、借用、歸還、電子郵件通知（含確認連結）
// =============================================

// ⚠️⚠️⚠️ 請替換成你的實際 Sheet ID ⚠️⚠️⚠️
const SPREADSHEET_ID = '1zW8SfCm8YtKwSfEnxqACn78TJaY4XIY5YL-OPZHliGY';

// 設定：工作表名稱
const SHEET_NAME = '工作表 1';
const KEEPER_SHEET_NAME = 'Keeper 聯絡資訊';
const HISTORY_SHEET_NAME = '歷史紀錄';
const AVATAR_SHEET_NAME = '頭像資料';

// 頭像資料夾 ID（請替換成你的 Google Drive 頭像資料夾 ID）
// 建立方式：在 Google Drive 建立一個資料夾，分享為「知道連結的使用者」可檢視，然後複製資料夾網址的最後一段
const AVATAR_FOLDER_ID = '15vkYY7wO1HyNKa0aruqDiLSDyKWuS1af';

// 欄位索引對照（0-indexed）
const COLS = {
  fix_type: 0,     // A 欄
  fix_no: 1,       // B 欄
  device_name: 2,  // C 欄
  qty_asset: 3,    // D 欄
  keeper: 4,       // E 欄
  status: 5,       // F 欄
  borrower: 6,     // G 欄
  dt_borrow: 7,    // H 欄
  dt_due: 8,       // I 欄
  dt_return: 9,    // J 欄
  return_confirmed: 10  // K 欄
};

// 電子郵件設定
const EMAIL_CONFIG = {
  enabled: true,
  subject_prefix: '[MT 設備系統]',
  borrow_subject: '設備借用通知',
  return_subject: '設備歸還通知',
  return_confirm_subject: '歸還確認通知',
  web_app_url: ' https://lobbster1234-ai.github.io/mt-equipment-system/'
};

/**
 * GET 請求處理
 */
function doGet(e) {
  try {
    const action = e.parameter.action || 'query';
    
    if (action === 'query') {
      return queryEquipment(e.parameter);
    } else if (action === 'register') {
      return registerEquipment({
        fix_type: e.parameter.fix_type,
        fix_no: e.parameter.fix_no,
        device_name: e.parameter.device_name,
        qty_asset: e.parameter.qty_asset,
        keeper: e.parameter.keeper
      });
    } else if (action === 'return') {
      return returnEquipment({
        fix_no: e.parameter.fix_no,
        dt_return: e.parameter.dt_return
      });
    } else if (action === 'confirmReturn') {
      return confirmReturn({
        fix_no: e.parameter.fix_no,
        keeper_email: e.parameter.keeper_email
      });
    } else if (action === 'loginAdmin') {
      return loginAdmin(e.parameter);
    } else if (action === 'borrow') {
      return borrowEquipment({
        fix_no: e.parameter.fix_no,
        borrower: e.parameter.borrower,
        dt_borrow: e.parameter.dt_borrow,
        dt_due: e.parameter.dt_due
      });
    } else if (action === 'getEquipmentInfo') {
      return getEquipmentInfo(e.parameter.fix_no);
    } else if (action === 'history') {
      return queryHistory(e.parameter);
    } else if (action === 'uploadAvatar') {
      return uploadAvatar({
        user_name: e.parameter.user_name,
        image_data: e.parameter.image_data,
        file_name: e.parameter.file_name
      });
    } else if (action === 'getAvatarList') {
      return getAvatarList();
    } else if (action === 'test') {
      return successResponse({
        status: 'ok',
        message: 'GAS 連線成功！',
        timestamp: new Date().toISOString()
      });
    }
    
    return errorResponse('未知的 action: ' + action);
  } catch (err) {
    return errorResponse(err.message);
  }
}

/**
 * 處理 CORS preflight 請求
 */
function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST 請求處理
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === 'query') {
      return queryEquipment(data);
    } else if (action === 'register') {
      return registerEquipment(data);
    } else if (action === 'uploadAvatar') {
      return uploadAvatar(data);
    } else if (action === 'test') {
      return successResponse({
        status: 'ok',
        message: 'GAS POST 連線成功！',
        timestamp: new Date().toISOString()
      });
    }
    
    return errorResponse('未知的 action: ' + action);
  } catch (err) {
    return errorResponse(err.message);
  }
}

/**
 * 查詢設備
 */
function queryEquipment(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return errorResponse(`找不到工作表：${SHEET_NAME}`);
  }
  
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  
  const keyword = (params.keyword || '').toLowerCase();
  const status = params.status || '';
  
  const filtered = rows.filter((row) => {
    if (!row[COLS.fix_no] && !row[COLS.device_name]) return false;
    
    if (keyword) {
      const fixNo = (row[COLS.fix_no] || '').toString().toLowerCase();
      const deviceName = (row[COLS.device_name] || '').toString().toLowerCase();
      const keeper = (row[COLS.keeper] || '').toString().toLowerCase();
      const borrower = (row[COLS.borrower] || '').toString().toLowerCase();
      
      if (!fixNo.includes(keyword) && !deviceName.includes(keyword) && !keeper.includes(keyword) && !borrower.includes(keyword)) {
        return false;
      }
    }
    
    if (status) {
      const rowStatus = (row[COLS.status] || '').toString().trim().toLowerCase();
      const filterStatus = status.toString().trim().toLowerCase();
      
      // 可借用：匹配 'available'、'可借用'、或空值
      if (filterStatus === 'available') {
        if (rowStatus !== 'available' && rowStatus !== '可借用' && rowStatus !== '') {
          return false;
        }
      }
      // 已借出：匹配 'borrowed' 或 '已借出'
      else if (filterStatus === 'borrowed') {
        if (rowStatus !== 'borrowed' && rowStatus !== '已借出') {
          return false;
        }
      }
      // 其他狀態：精確匹配
      else if (rowStatus !== filterStatus) {
        return false;
      }
    }
    
    return true;
  });
  
  const result = filtered.map(row => ({
    fix_type: row[COLS.fix_type] || '',
    fix_no: row[COLS.fix_no] || '',
    device_name: row[COLS.device_name] || '',
    qty_asset: row[COLS.qty_asset] || '1',
    keeper: row[COLS.keeper] || '',
    status: row[COLS.status] || 'available',
    borrower: row[COLS.borrower] || '',
    dt_borrow: formatDate(row[COLS.dt_borrow]),
    dt_due: formatDate(row[COLS.dt_due]),
    dt_return: formatDate(row[COLS.dt_return]),
    return_confirmed: row[COLS.return_confirmed] || false
  }));
  
  return successResponse(result);
}

/**
 * 登記設備
 */
function registerEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return errorResponse(`找不到工作表：${SHEET_NAME}`);
  }
  
  const newRow = [
    data.fix_type || '',
    data.fix_no || '',
    data.device_name || '',
    data.qty_asset || '1',
    data.keeper || '',
    'available',
    '',
    '',
    '',
    '',
    false
  ];
  
  sheet.appendRow(newRow);
  
  return successResponse({
    success: true,
    message: '設備登記成功',
    fix_no: data.fix_no
  });
}

/**
 * 借用設備
 */
function borrowEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return errorResponse(`找不到工作表：${SHEET_NAME}`);
  }
  
  const fixNo = data.fix_no;
  const borrower = data.borrower;
  const dtBorrow = data.dt_borrow || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const dtDue = data.dt_due || '';
  
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  const borrowerCol = COLS.borrower;
  const dtBorrowCol = COLS.dt_borrow;
  const dtDueCol = COLS.dt_due;
  const dtReturnCol = COLS.dt_return;
  const keeperCol = COLS.keeper;
  const deviceNameCol = COLS.device_name;
  
  let foundRow = -1;
  const lastRow = sheet.getLastRow();
  
  for (let i = 2; i <= lastRow; i++) {
    const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
    if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
      foundRow = i;
      break;
    }
  }
  
  if (foundRow === -1) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const currentStatus = sheet.getRange(foundRow, statusCol + 1).getValue();
  if (currentStatus === 'borrowed') {
    return errorResponse('設備已經借出');
  }
  
  sheet.getRange(foundRow, statusCol + 1).setValue('borrowed');
  sheet.getRange(foundRow, borrowerCol + 1).setValue(borrower);
  sheet.getRange(foundRow, dtBorrowCol + 1).setValue(dtBorrow);
  sheet.getRange(foundRow, dtDueCol + 1).setValue(dtDue);
  sheet.getRange(foundRow, dtReturnCol + 1).setValue(''); // 清除歸還日期
  sheet.getRange(foundRow, COLS.return_confirmed + 1).setValue(false);
  
  const keeper = sheet.getRange(foundRow, keeperCol + 1).getValue();
  const deviceName = sheet.getRange(foundRow, deviceNameCol + 1).getValue();
  
  // 記錄歷史紀錄
  logHistory('borrow', fixNo, deviceName, borrower, keeper, dtBorrow, dtDue, '');
  
  if (EMAIL_CONFIG.enabled && keeper) {
    sendBorrowEmail(keeper, fixNo, deviceName, borrower, dtBorrow, dtDue);
  }
  
  return successResponse({
    message: '借用成功',
    fix_no: fixNo,
    borrower: borrower,
    dt_borrow: dtBorrow,
    dt_due: dtDue
  });
}

/**
 * 歸還設備
 */
function returnEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return errorResponse(`找不到工作表：${SHEET_NAME}`);
  }
  
  const fixNo = data.fix_no;
  // 強制使用當天台北時間，忽略前端傳來的日期
  const dtReturn = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  const dtReturnCol = COLS.dt_return;
  const keeperCol = COLS.keeper;
  const deviceNameCol = COLS.device_name;
  const borrowerCol = COLS.borrower;
  const dtBorrowCol = COLS.dt_borrow;
  const dtDueCol = COLS.dt_due;
  
  let foundRow = -1;
  const lastRow = sheet.getLastRow();
  
  for (let i = 2; i <= lastRow; i++) {
    const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
    if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
      foundRow = i;
      break;
    }
  }
  
  if (foundRow === -1) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const currentStatus = (sheet.getRange(foundRow, statusCol + 1).getValue() || '').toString().trim().toLowerCase();
  const isBorrowed = currentStatus === 'borrowed' || currentStatus === '借用中' || currentStatus === '已借出' || currentStatus === '使用中';
  
  if (!isBorrowed) {
    return errorResponse(`設備狀態不是借用中（當前狀態：${currentStatus || '空'}）`);
  }
  
  const keeper = sheet.getRange(foundRow, keeperCol + 1).getValue();
  const deviceName = sheet.getRange(foundRow, deviceNameCol + 1).getValue();
  const borrower = sheet.getRange(foundRow, borrowerCol + 1).getValue();
  const dtBorrowVal = sheet.getRange(foundRow, dtBorrowCol + 1).getValue();
  const dtDueVal = sheet.getRange(foundRow, dtDueCol + 1).getValue();
  
  sheet.getRange(foundRow, dtReturnCol + 1).setValue(dtReturn);
  // 更新狀態為「歸還認證中」，避免重複按下歸還按鈕
  sheet.getRange(foundRow, statusCol + 1).setValue('return_pending');
  
  // 記錄歷史紀錄（歸還動作：借用日期、預計歸還、實際歸還日期）
  logHistory('return', fixNo, deviceName, borrower, keeper, dtBorrowVal, dtDueVal, dtReturn);
  
  if (EMAIL_CONFIG.enabled && keeper) {
    sendReturnEmail(keeper, fixNo, deviceName, borrower, dtReturn);
  }
  
  return successResponse({
    message: '歸還通知已發送，請等待 Keeper 確認',
    fix_no: fixNo,
    dt_return: dtReturn
  });
}

/**
 * 取得設備資訊（用於確認頁面）
 */
function getEquipmentInfo(fixNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return errorResponse(`找不到工作表：${SHEET_NAME}`);
  }
  
  const fixNoCol = COLS.fix_no;
  const keeperCol = COLS.keeper;
  
  let foundRow = -1;
  const lastRow = sheet.getLastRow();
  
  for (let i = 2; i <= lastRow; i++) {
    const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
    if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
      foundRow = i;
      break;
    }
  }
  
  if (foundRow === -1) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const row = sheet.getRange(foundRow, 1, 1, 11).getValues()[0];
  
  return successResponse({
    fix_type: row[COLS.fix_type] || '',
    fix_no: row[COLS.fix_no] || '',
    device_name: row[COLS.device_name] || '',
    qty_asset: row[COLS.qty_asset] || '1',
    keeper: row[COLS.keeper] || '',
    status: row[COLS.status] || 'available',
    borrower: row[COLS.borrower] || '',
    dt_borrow: forceFormatDate(row[COLS.dt_borrow]),
    dt_due: forceFormatDate(row[COLS.dt_due]),
    dt_return: forceFormatDate(row[COLS.dt_return]),
    return_confirmed: row[COLS.return_confirmed] || false
  });
}

/**
 * 確認歸還（Keeper 點擊確認連結）
 */
function confirmReturn(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return errorResponse(`找不到工作表：${SHEET_NAME}`);
  }
  
  const fixNo = data.fix_no;
  // 不再需要 keeper_email 驗證，因為 token 已經由系統生成並發送給正確的 Keeper
  
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  const borrowerCol = COLS.borrower;
  const dtBorrowCol = COLS.dt_borrow;
  const dtDueCol = COLS.dt_due;
  const dtReturnCol = COLS.dt_return;
  const returnConfirmedCol = COLS.return_confirmed;
  const keeperCol = COLS.keeper;
  const deviceNameCol = COLS.device_name;
  
  let foundRow = -1;
  const lastRow = sheet.getLastRow();
  
  for (let i = 2; i <= lastRow; i++) {
    const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
    if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
      foundRow = i;
      break;
    }
  }
  
  if (foundRow === -1) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const keeperName = sheet.getRange(foundRow, keeperCol + 1).getValue();
  const deviceName = sheet.getRange(foundRow, deviceNameCol + 1).getValue();
  const borrower = sheet.getRange(foundRow, borrowerCol + 1).getValue();
  const dtBorrowVal = sheet.getRange(foundRow, dtBorrowCol + 1).getValue();
  const dtDueVal = sheet.getRange(foundRow, dtDueCol + 1).getValue();
  const dtReturnVal = sheet.getRange(foundRow, dtReturnCol + 1).getValue();
  
  // 直接確認歸還，不再驗證 email
  Logger.log(`確認歸還：${fixNo}，保管人：${keeperName}`);
  
  sheet.getRange(foundRow, statusCol + 1).setValue('available');
  sheet.getRange(foundRow, returnConfirmedCol + 1).setValue(true);
  sheet.getRange(foundRow, borrowerCol + 1).setValue('');
  sheet.getRange(foundRow, dtBorrowCol + 1).setValue('');
  sheet.getRange(foundRow, dtDueCol + 1).setValue('');
  
  // 記錄歷史紀錄（確認歸還）
  logHistory('confirm', fixNo, deviceName, borrower || '', keeperName, dtBorrowVal || '', dtDueVal || '', dtReturnVal || '');
  
  if (EMAIL_CONFIG.enabled && keeperName) {
    sendReturnConfirmEmail(keeperName, fixNo, deviceName);
  }
  
  return successResponse({
    message: '歸還已確認，設備狀態已更新為可借用',
    fix_no: fixNo
  });
}

/**
 * 發送借用通知郵件
 */
function sendBorrowEmail(keeper, fixNo, deviceName, borrower, dtBorrow, dtDue) {
  try {
    const keeperEmail = getKeeperEmail(keeper);
    
    Logger.log(`發送借用通知 - keeper: ${keeper}, email: ${keeperEmail}`);
    
    if (!keeperEmail) {
      Logger.log(`找不到 ${keeper} 的電子郵件，跳過寄信`);
      return;
    }
    
    const subject = `${EMAIL_CONFIG.subject_prefix} ${EMAIL_CONFIG.borrow_subject}`;
    const body = `親愛的 ${keeper} 您好：

有人借用了您保管的設備，詳情如下：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
👤 借用人：${borrower}
📅 借用日期：${dtBorrow}
⏰ 預計歸還：${dtDue || '未設定'}

請留意設備歸還狀況。

---
MT 部門設備管理系統 自動通知`.trim();
    
    MailApp.sendEmail(keeperEmail, subject, body);
    Logger.log(`已發送借用通知給 ${keeperEmail}`);
  } catch (err) {
    Logger.error('發送借用通知郵件失敗:', err);
  }
}

/**
 * 發送歸還通知郵件（包含確認連結）
 */
function sendReturnEmail(keeper, fixNo, deviceName, borrower, dtReturn) {
  try {
    const keeperEmail = getKeeperEmail(keeper);
    
    if (!keeperEmail) {
      Logger.log(`找不到 ${keeper} 的電子郵件，跳過寄信`);
      return;
    }
    
    // 建立確認連結（包含 fix_no 和 keeper_email）
    const token = Utilities.base64Encode(`${fixNo}:${keeperEmail}:${Date.now()}`);
    const confirmUrl = `${EMAIL_CONFIG.web_app_url}/confirm.html?token=${encodeURIComponent(token)}`;
    
    const subject = `${EMAIL_CONFIG.subject_prefix} ${EMAIL_CONFIG.return_subject}`;
    const body = `親愛的 ${keeper} 您好：

您保管的設備已被歸還，請確認收到：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
👤 原借用人：${borrower}
📅 歸還日期：${dtReturn}

✅ 請點擊下方按鈕確認歸還：
${confirmUrl}

或者複製以下網址到瀏覽器開啟：
${confirmUrl}

---
MT 部門設備管理系統 自動通知`.trim();
    
    MailApp.sendEmail(keeperEmail, subject, body);
    Logger.log(`已發送歸還通知給 ${keeperEmail}`);
  } catch (err) {
    Logger.error('發送歸還通知郵件失敗:', err);
  }
}

/**
 * 發送歸還確認郵件
 */
function sendReturnConfirmEmail(keeper, fixNo, deviceName) {
  try {
    const keeperEmail = getKeeperEmail(keeper);
    
    if (!keeperEmail) {
      return;
    }
    
    const subject = `${EMAIL_CONFIG.subject_prefix} ${EMAIL_CONFIG.return_confirm_subject}`;
    const body = `親愛的 ${keeper} 您好：

您已確認收到歸還的設備：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}

✅ 設備狀態已更新為「可借用」

感謝您的配合！

---
MT 部門設備管理系統 自動通知`.trim();
    
    MailApp.sendEmail(keeperEmail, subject, body);
  } catch (err) {
    Logger.error('發送歸還確認郵件失敗:', err);
  }
}

/**
 * 取得 Keeper 的電子郵件地址
 */
function getKeeperEmail(keeperName) {
  if (keeperName && keeperName.includes('@')) {
    Logger.log(`keeper 欄位直接是 email: ${keeperName}`);
    return keeperName;
  }
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const keeperSheet = ss.getSheetByName(KEEPER_SHEET_NAME);
    
    if (!keeperSheet) {
      Logger.log(`找不到工作表：${KEEPER_SHEET_NAME}`);
      return null;
    }
    
    const data = keeperSheet.getDataRange().getValues();
    Logger.log(`Keeper 聯絡資訊工作表共有 ${data.length} 列`);
    Logger.log(`查找的 keeper 姓名：「${keeperName}」`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0] ? row[0].toString().trim() : '';
      const email = row[1] ? row[1].toString().trim() : '';
      
      Logger.log(`比對第 ${i+1} 列：姓名「${name}」, email「${email}」`);
      
      if (name && name === keeperName) {
        Logger.log(`找到匹配的 email: ${email}`);
        return email;
      }
    }
    
    Logger.log(`在 ${KEEPER_SHEET_NAME} 中找不到 ${keeperName} 的 email`);
    return null;
  } catch (err) {
    Logger.error('讀取 Keeper 聯絡資訊失敗:', err);
    return null;
  }
}

/**
 * 取得頭像圖片 URL
 */
function getAvatarUrl(userName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let avatarSheet = ss.getSheetByName(AVATAR_SHEET_NAME);
    
    // 如果頭像工作表不存在，返回 null
    if (!avatarSheet) {
      return null;
    }
    
    const data = avatarSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = (row[0] || '').toString().trim();
      const avatarUrl = (row[1] || '').toString().trim();
      
      if (name === userName.trim() && avatarUrl) {
        return avatarUrl;
      }
    }
    
    return null;
  } catch (err) {
    Logger.log('取得頭像失敗: ' + err.message);
    return null;
  }
}

/**
 * 上傳頭像圖片（存到 Sheet，不是 Drive）
 */
function uploadAvatar(data) {
  try {
    Logger.log('uploadAvatar 收到參數，user_name: ' + data.user_name);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userName = data.user_name;
    const imageData = data.image_data; // base64 編碼的圖片（含 data:image/jpeg;base64, 前綴）
    
    if (!userName || !imageData) {
      return errorResponse('缺少必要參數');
    }
    
    // 儲存到工作表
    let avatarSheet = ss.getSheetByName(AVATAR_SHEET_NAME);
    if (!avatarSheet) {
      // 建立頭像工作表
      avatarSheet = ss.insertSheet(AVATAR_SHEET_NAME);
      avatarSheet.appendRow(['姓名', '頭像Base64', '更新時間']);
    }
    
    // 檢查是否已有記錄
    const dataRange = avatarSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < dataRange.length; i++) {
      if ((dataRange[i][0] || '').toString().trim() === userName.trim()) {
        // 更新現有記錄
        avatarSheet.getRange(i + 1, 2).setValue(imageData);
        avatarSheet.getRange(i + 1, 3).setValue(new Date());
        found = true;
        break;
      }
    }
    
    if (!found) {
      // 新增記錄
      avatarSheet.appendRow([userName, imageData, new Date()]);
    }
    
    Logger.log('頭像儲存成功: ' + userName);
    
    return successResponse({
      success: true,
      message: '頭像上傳成功',
      url: imageData  // 直接回傳 base64 data URL
    });
  } catch (err) {
    Logger.log('頭像上傳失敗: ' + err.message);
    return errorResponse('頭像上傳失敗: ' + err.message);
  }
}


/**
 * 取得所有頭像列表
 */
function getAvatarList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let avatarSheet = ss.getSheetByName(AVATAR_SHEET_NAME);
    
    if (!avatarSheet) {
      return successResponse([]);
    }
    
    const data = avatarSheet.getDataRange().getValues();
    const result = [];
    
    // 跳過標題列
    for (let i = 1; i < data.length; i++) {
      const name = (data[i][0] || '').toString().trim();
      const avatarUrl = (data[i][1] || '').toString().trim();
      
      if (name && avatarUrl) {
        result.push({
          name: name,
          avatar_url: avatarUrl
        });
      }
    }
    
    Logger.log('取得頭像列表：' + result.length + ' 個');
    return successResponse(result);
  } catch (err) {
    Logger.log('取得頭像列表失敗：' + err.message);
    return errorResponse('取得頭像列表失敗：' + err.message);
  }
}
/**
 * 輔助函式：成功回應
 */
function successResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 輔助函式：錯誤回應
 */
function errorResponse(message) {
  return ContentService.createTextOutput(JSON.stringify({
    error: message,
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 輔助函式：格式化日期為 yyyy-MM-dd 字串（台北時區）
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  // 如果是 Date 物件，使用 Utilities.formatDate 格式化为台北时区
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  
  // 如果是字串，檢查是否為 yyyy-MM-dd 格式
  if (typeof dateValue === 'string') {
    const trimmed = dateValue.trim();
    // 如果已經是 yyyy-MM-dd 格式，直接返回
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }
    // 嘗試解析為 Date
    if (trimmed) {
      const date = new Date(trimmed);
      if (!isNaN(date.getTime())) {
        return Utilities.formatDate(date, 'Asia/Taipei', 'yyyy-MM-dd');
      }
      return trimmed;
    }
    return '';
  }
  
  // 如果是數字（時間戳），轉換為 Date
  if (typeof dateValue === 'number') {
    return Utilities.formatDate(new Date(dateValue), 'Asia/Taipei', 'yyyy-MM-dd');
  }
  
  return String(dateValue);
}

/**
 * 記錄歷史紀錄
 */
function logHistory(action, fixNo, deviceName, borrower, keeper, dtAction, dtDue, dtConfirmed) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    
    // 如果歷史紀錄工作表不存在，建立它
    if (!sheet) {
      sheet = ss.insertSheet(HISTORY_SHEET_NAME);
      // 建立標題列
      sheet.appendRow(['時間戳', '動作', '設備編號', '設備名稱', '借用人', '保管人', '借用日期', '預計歸還', '實際歸還/確認日期']);
    }
    
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
    
    sheet.appendRow([
      now,
      action,
      fixNo,
      deviceName,
      borrower || '',
      keeper || '',
      dtAction || '',
      dtDue || '',
      dtConfirmed || ''
    ]);
    
    Logger.log(`已記錄歷史紀錄：${action} - ${fixNo}`);
  } catch (err) {
    Logger.error('記錄歷史紀錄失敗:', err);
  }
}

/**
 * 強制格式化日期為 yyyy-MM-dd（不使用 Utilities.formatDate）
 */
function forceFormatDate(value) {
  if (!value) return '';
  
  let dateObj;
  
  // 如果是 Date 物件
  if (value instanceof Date) {
    dateObj = value;
  }
  // 如果是數字（時間戳）
  else if (typeof value === 'number') {
    dateObj = new Date(value);
  }
  // 如果是字串
  else if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return '';
    
    // 如果已經是 yyyy-MM-dd 格式，直接返回
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }
    
    // 嘗試解析為 Date
    dateObj = new Date(trimmed);
  }
  else {
    return String(value).trim();
  }
  
  // 格式化為 yyyy-MM-dd
  if (!isNaN(dateObj.getTime())) {
    // 使用台北時區（GMT+8）
    const taipeiOffset = 8 * 60; // 8小時轉分鐘
    const localOffset = dateObj.getTimezoneOffset(); // 本地時區偏移（分鐘）
    const offsetDiff = taipeiOffset + localOffset; // 差異
    const taipeiDate = new Date(dateObj.getTime() + offsetDiff * 60000);
    
    const year = taipeiDate.getFullYear();
    const month = String(taipeiDate.getMonth() + 1).padStart(2, '0');
    const day = String(taipeiDate.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  
  return '';
}

/**
 * 格式化日期顯示為 yyyy-MM-dd
 */
function formatDisplayDate(value) {
  return forceFormatDate(value);
}

/**
 * 解析歷史紀錄中的日期（處理各種格式）
 */
function parseHistoryDate(value) {
  if (!value) return '';
  
  // 如果已經是 yyyy-MM-dd 格式，直接返回
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
    return value.trim();
  }
  
  // 如果是 Date 物件
  if (value instanceof Date) {
    if (!isNaN(value.getTime())) {
      return Utilities.formatDate(value, 'Asia/Taipei', 'yyyy-MM-dd');
    }
    return '';
  }
  
  // 如果是數字（時間戳）
  if (typeof value === 'number') {
    try {
      return Utilities.formatDate(new Date(value), 'Asia/Taipei', 'yyyy-MM-dd');
    } catch (e) {
      return '';
    }
  }
  
  // 如果是字串，嘗試解析
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return '';
    
    // 嘗試解析常見日期格式
    try {
      const date = new Date(trimmed);
      if (!isNaN(date.getTime())) {
        return Utilities.formatDate(date, 'Asia/Taipei', 'yyyy-MM-dd');
      }
    } catch (e) {
      // 解析失敗，返回原始字串（如果看起來像日期）
      if (/\d{4}[-/]\d{2}[-/]\d{2}/.test(trimmed)) {
        return trimmed;
      }
    }
  }
  
  return '';
}

/**
 * 查詢歷史紀錄
 */
function queryHistory(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  
  if (!sheet) {
    return successResponse([]);
  }
  
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  
  const keyword = (params.keyword || '').toLowerCase();
  const actionType = params.actionType || '';  // 修正：使用 actionType 避免與 action='history' 衝突
  
  const filtered = rows.filter((row) => {
    if (!row[1]) return false;
    
    if (keyword) {
      const fixNo = (row[2] || '').toString().toLowerCase();
      const deviceName = (row[3] || '').toString().toLowerCase();
      const borrower = (row[4] || '').toString().toLowerCase();
      const keeper = (row[5] || '').toString().toLowerCase();
      
      if (!fixNo.includes(keyword) && !deviceName.includes(keyword) && !borrower.includes(keyword) && !keeper.includes(keyword)) {
        return false;
      }
    }
    
    if (actionType && (row[1] || '').toString() !== actionType) {
      return false;
    }
    
    return true;
  });
  
  filtered.sort((a, b) => {
    const aTime = new Date(a[0]);
    const bTime = new Date(b[0]);
    return bTime - aTime;
  });
  
  const result = filtered.map(row => {
    const action = row[1] || '';
    
    // 根據動作類型，正確對應日期欄位
    // 欄位索引：0=時間戳，1=動作，2=設備編號，3=設備名稱，4=借用人，5=保管人，6=借用日期，7=預計歸還，8=歸還日期
    let dt_borrow = '';
    let dt_due = '';
    let dt_return = '';
    let return_confirmed = false;
    
    if (action === 'borrow') {
      // 借用：row[6]=借用日期，row[7]=預計歸還，row[8]=空
      dt_borrow = formatDisplayDate(row[6]);
      dt_due = formatDisplayDate(row[7]);
      dt_return = '';
      return_confirmed = false;
    } else if (action === 'return') {
      // 歸還：row[6]=借用日期，row[7]=預計歸還，row[8]=歸還日期
      dt_borrow = formatDisplayDate(row[6]);
      dt_due = formatDisplayDate(row[7]);
      dt_return = formatDisplayDate(row[8]);
      return_confirmed = false;
    } else if (action === 'confirm') {
      // 確認：row[6]=借用日期，row[7]=預計歸還，row[8]=確認日期
      dt_borrow = formatDisplayDate(row[6]);
      dt_due = formatDisplayDate(row[7]);
      dt_return = formatDisplayDate(row[8]);
      return_confirmed = true;
    }
    
    return {
      timestamp: row[0] || '',
      action: action,
      fix_no: row[2] || '',
      device_name: row[3] || '',
      borrower: row[4] || '',
      keeper: row[5] || '',
      dt_borrow: dt_borrow,
      dt_due: dt_due,
      dt_return: dt_return,
      return_confirmed: return_confirmed
    };
  });
  
  return successResponse(result);
}

/**
 * 管理員登入驗證
 */
function loginAdmin(data) {
  const email = data.email;
  const password = data.password;
  
  if (!email || !password) {
    return errorResponse('請提供電子郵件和密碼');
  }
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const keeperSheet = ss.getSheetByName(KEEPER_SHEET_NAME);
    
    if (!keeperSheet) {
      return errorResponse('找不到 Keeper 聯絡資訊工作表');
    }
    
    const sheetData = keeperSheet.getDataRange().getValues();
    Logger.log('Keeper 工作表資料筆數：' + sheetData.length);
    
    // 從第 2 列開始（跳過標題）
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const rowName = (row[0] || '').toString().trim();      // A 欄 - 姓名
      const rowEmail = (row[1] || '').toString().trim();      // B 欄 - 電子郵件
      const rowAccount = (row[2] || '').toString().trim();   // C 欄 - 帳號
      const rowPassword = (row[3] || '').toString().trim();  // D 欄 - 密碼
      
      Logger.log('檢查第 ' + (i+1) + ' 列：email=' + rowEmail + ', account=' + rowAccount + ', name=' + rowName);
      
      // 檢查電子郵件或帳號是否匹配
      if (rowEmail === email || rowAccount === email) {
        Logger.log('找到匹配：' + email);
        
        // 找到匹配的電子郵件，檢查密碼
        if (!rowPassword) {
          Logger.log('密碼為空');
          return errorResponse('此帳號尚未設定密碼，請聯繫管理員');
        }
        
        if (rowPassword === password) {
          Logger.log('登入成功：' + rowName);
          // 登入成功
          return successResponse({
            name: rowName,
            email: rowEmail,
            role: 'admin'
          });
        } else {
          Logger.log('密碼錯誤');
          // 密碼錯誤
          return errorResponse('密碼錯誤');
        }
      }
    }
    
    Logger.log('找不到帳號：' + email);
    return errorResponse('找不到此管理員帳號');
    
  } catch (err) {
    Logger.error('登入失敗:', err);
    return errorResponse('登入失敗：' + err.message);
  }
}
