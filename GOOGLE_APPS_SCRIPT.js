// =============================================
// MT 設備系統 - Google Apps Script 後端（完整版）
// =============================================
// 功能：查詢、登記、借用、歸還、借用審核、電子郵件通知（含確認連結）
// =============================================

// 新增待審核借用工作表名稱

// ⚠️⚠️⚠️ 請替換成你的實際 Sheet ID ⚠️⚠️⚠️
const SPREADSHEET_ID = '1zW8SfCm8YtKwSfEnxqACn78TJaY4XIY5YL-OPZHliGY';

// 設定：工作表名稱
const SHEET_NAME = '工作表 1';           // 主要設備清單（部門助理年度更新）
const BORROW_REQUEST_SHEET_NAME = '借用申請';
const SHEET_NAME_WEB = '網站新增設備';    // 網站新增的設備（管理員手動新增）
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
  web_app_url: 'https://lobbster1234-ai.github.io/mt-equipment-system/'
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
      const params = {
        email: e.parameter.email,
        password: e.parameter.password
      };
      Logger.log('loginAdmin 參數：' + JSON.stringify(params));
      return loginAdmin(params);
    } else if (action === 'setupPassword') {
      return setupPassword({
        email: e.parameter.email,
        newPassword: e.parameter.newPassword
      });
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
    } else if (action === 'updateEquipment') {
      return updateEquipment({
        fix_no: e.parameter.fix_no,
        device_name: e.parameter.device_name,
        fix_type: e.parameter.fix_type,
        qty_asset: e.parameter.qty_asset
      });
    } else if (action === 'deleteEquipment') {
      return deleteEquipment({
        fix_no: e.parameter.fix_no
      });
    } else if (action === 'requestBorrow') {
      const requestData = {
        fix_no: e.parameter.fix_no || '',
        borrower: e.parameter.borrower || '',
        borrower_email: e.parameter.borrower_email || '',
        dt_borrow: e.parameter.dt_borrow || '',
        dt_due: e.parameter.dt_due || ''
      };
      Logger.log('requestBorrow 接收到的参数: ' + JSON.stringify(requestData));
      return requestBorrow(requestData);
    } else if (action === 'approveBorrow') {
      return approveBorrow({
        request_id: e.parameter.request_id
      }, e);
    } else if (action === 'rejectBorrow') {
      return rejectBorrow({
        request_id: e.parameter.request_id
      }, e);
    } else if (action === 'getBorrowRequest') {
      return getBorrowRequest({
        request_id: e.parameter.request_id
      }, e);
    } else if (action === 'getEmailByName') {
      return getEmailByName({
        name: e.parameter.name
      });
    } else if (action === 'postponeDueDate') {
      return postponeDueDate({
        fix_no: e.parameter.fix_no,
        new_due_date: e.parameter.new_due_date
      });
    } else if (action === 'requestPostpone') {
      return requestPostpone({
        fix_no: e.parameter.fix_no,
        new_due_date: e.parameter.new_due_date
      });
    } else if (action === 'getPostponeRequest') {
      return getPostponeRequest({
        request_id: e.parameter.request_id
      }, e);
    } else if (action === 'approvePostpone') {
      return approvePostpone({
        request_id: e.parameter.request_id
      }, e);
    } else if (action === 'rejectPostpone') {
      return rejectPostpone({
        request_id: e.parameter.request_id
      }, e);
    } else if (action === 'deptBorrow') {
      return deptBorrow({
        device_name: e.parameter.device_name,
        borrower: e.parameter.borrower,
        borrower_email: e.parameter.borrower_email,
        dt_borrow: e.parameter.dt_borrow,
        dt_due: e.parameter.dt_due
      });
    } else if (action === 'deptReturn') {
      return deptReturn({
        id: e.parameter.id
      });
    } else if (action === 'getDeptBorrowList') {
      return getDeptBorrowList();
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
    } else if (action === 'loginAdmin') {
      return loginAdmin(data);
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
  
  // 同時讀取兩個工作表
  const sheet1 = ss.getSheetByName(SHEET_NAME);
  const sheet2 = ss.getSheetByName(SHEET_NAME_WEB);
  
  if (!sheet1) {
    return errorResponse(`找不到工作表：${SHEET_NAME}`);
  }
  
  // 合併兩個工作表的資料
  let allData = [];
  
  // 讀取工作表 1（主要清單）
  const data1 = sheet1.getDataRange().getValues();
  allData = allData.concat(data1.slice(1));  // 跳過標題列
  
  // 讀取工作表 2（網站新增設備），如果存在的話
  if (sheet2) {
    const data2 = sheet2.getDataRange().getValues();
    allData = allData.concat(data2.slice(1));  // 跳過標題列
    Logger.log('合併網站新增設備：' + (data2.length - 1) + '筆');
  }
  
  const keyword = (params.keyword || '').toLowerCase();
  const status = params.status || '';
  
  const filtered = allData.filter((row) => {
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
  
  Logger.log('查詢結果：共 ' + result.length + ' 筆設備');
  return successResponse(result);
}

/**
 * 登記設備
 */
function registerEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // 網站新增的設備寫入「網站新增設備」工作表
  const sheet = ss.getSheetByName(SHEET_NAME_WEB);
  
  if (!sheet) {
    return errorResponse(`找不到工作表：${SHEET_NAME_WEB}，請先建立此工作表`);
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
    message: '設備登記成功（已存入網站新增設備工作表）',
    fix_no: data.fix_no
  });
}

/**
 * 借用設備 - 支援兩個工作表
 */
function borrowEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
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
  
  // 先在「工作表 1」查找
  let sheet = ss.getSheetByName(SHEET_NAME);
  let foundRow = -1;
  let targetSheet = null;
  let sheetSource = '';
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
        foundRow = i;
        targetSheet = sheet;
        sheetSource = SHEET_NAME;
        break;
      }
    }
  }
  
  // 如果「工作表 1」找不到，在「網站新增設備」查找
  if (foundRow === -1) {
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          sheetSource = SHEET_NAME_WEB;
          break;
        }
      }
    }
  }
  
  if (foundRow === -1 || !targetSheet) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  Logger.log(`找到設備於工作表: ${targetSheet.getName()}, 行號: ${foundRow}`);
  
  const currentStatus = targetSheet.getRange(foundRow, statusCol + 1).getValue();
  Logger.log(`當前狀態: ${currentStatus}`);
  
  if (currentStatus === 'borrowed') {
    return errorResponse('設備已經借出');
  }
  if (currentStatus === 'borrow_pending') {
    return errorResponse('設備正在借用審核中');
  }
  
  // 不立即借出，而是設定為「借用審核中」狀態
  targetSheet.getRange(foundRow, statusCol + 1).setValue('borrow_pending');
  targetSheet.getRange(foundRow, borrowerCol + 1).setValue(borrower);
  targetSheet.getRange(foundRow, dtBorrowCol + 1).setValue(dtBorrow);
  targetSheet.getRange(foundRow, dtDueCol + 1).setValue(dtDue);
  targetSheet.getRange(foundRow, dtReturnCol + 1).setValue('');
  targetSheet.getRange(foundRow, COLS.return_confirmed + 1).setValue(false);
  
  // 確認寫入成功
  const verifyStatus = targetSheet.getRange(foundRow, statusCol + 1).getValue();
  const verifyBorrower = targetSheet.getRange(foundRow, borrowerCol + 1).getValue();
  Logger.log(`寫入驗證 - 狀態: ${verifyStatus}, 借用人: ${verifyBorrower}`);
  
  const keeper = targetSheet.getRange(foundRow, keeperCol + 1).getValue();
  const deviceName = targetSheet.getRange(foundRow, deviceNameCol + 1).getValue();
  
  // 記錄歷史（標記為 pending，等待審核）
  logHistory('borrow_pending', fixNo, deviceName, borrower, keeper, dtBorrow, dtDue, '');
  Logger.log(`已寫入歷史紀錄: ${fixNo}`);
  
  // 建立借用請求記錄（統一流程，讓 approveBorrow 能找到）
  let borrowRequestSheet = ss.getSheetByName(BORROW_REQUEST_SHEET_NAME);
  if (!borrowRequestSheet) {
    borrowRequestSheet = ss.insertSheet(BORROW_REQUEST_SHEET_NAME);
    borrowRequestSheet.appendRow(['請求ID', '設備編號', '設備名稱', '借用人', '借用人Email', '借用日期', '預計歸還', '保管人', '狀態', '建立時間']);
  }
  
  const requestId = Utilities.getUuid();
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const borrowerEmail = getKeeperEmail(borrower) || ''; // 從 Keeper 聯絡資訊取得借用人 email
  
  borrowRequestSheet.appendRow([
    requestId,
    fixNo,
    deviceName,
    borrower,
    borrowerEmail,
    dtBorrow,
    dtDue,
    keeper,
    'pending',
    timestamp
  ]);
  
  Logger.log(`借用請求已記錄到 ${BORROW_REQUEST_SHEET_NAME}，requestId: ${requestId}`);
  
  // 發送借用審核郵件給 Keeper
  if (EMAIL_CONFIG.enabled && keeper) {
    sendBorrowApprovalEmail(keeper, fixNo, deviceName, borrower, borrowerEmail, dtBorrow, dtDue, requestId);
  }
  
  Logger.log(`借用審核請求已建立：${fixNo}，借用人：${borrower}，等待 Keeper ${keeper} 審核`);
  return successResponse({
    message: '借用申請已送出，等待 Keeper 審核',
    request_id: requestId,
    fix_no: fixNo,
    borrower: borrower,
    keeper: keeper
  });
}

/**
 * 訪客借用請求（需要 Keeper 審核）
 */
function requestBorrow(data) {
  Logger.log('requestBorrow 函数被调用，data: ' + JSON.stringify(data));
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const fixNo = data.fix_no;
  const borrower = data.borrower;
  const borrowerEmail = data.borrower_email;
  const dtBorrow = data.dt_borrow || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const dtDue = data.dt_due || '';
  
  Logger.log(`解析参数 - fixNo: ${fixNo}, borrower: ${borrower}, email: ${borrowerEmail}`);
  
  if (!fixNo) {
    return errorResponse('缺少設備編號參數');
  }
  
  const fixNoCol = COLS.fix_no;
  const keeperCol = COLS.keeper;
  const deviceNameCol = COLS.device_name;
  
  // 查找設備
  let sheet = ss.getSheetByName(SHEET_NAME);
  let foundRow = -1;
  let targetSheet = null;
  let sheetSource = '';
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
        foundRow = i;
        targetSheet = sheet;
        sheetSource = SHEET_NAME;
        break;
      }
    }
  }
  
  if (foundRow === -1) {
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          sheetSource = SHEET_NAME_WEB;
          break;
        }
      }
    }
  }
  
  if (foundRow === -1 || !targetSheet) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const currentStatus = targetSheet.getRange(foundRow, COLS.status + 1).getValue();
  if (currentStatus === 'borrowed') {
    return errorResponse('設備已經借出');
  }
  
  const keeper = targetSheet.getRange(foundRow, keeperCol + 1).getValue();
  const deviceName = targetSheet.getRange(foundRow, deviceNameCol + 1).getValue();
  
  // 更新設備狀態為借用審核中
  targetSheet.getRange(foundRow, COLS.status + 1).setValue('borrow_pending');
  targetSheet.getRange(foundRow, COLS.borrower + 1).setValue(borrower);
  targetSheet.getRange(foundRow, COLS.dt_borrow + 1).setValue(dtBorrow);
  targetSheet.getRange(foundRow, COLS.dt_due + 1).setValue(dtDue);
  Logger.log(`設備 ${fixNo} 狀態已更新為 borrow_pending`);
  
  // 記錄歷史
  logHistory('borrow_pending', fixNo, deviceName, borrower, keeper, dtBorrow, dtDue, '');
  
  // 建立借用請求記錄
  let borrowRequestSheet = ss.getSheetByName(BORROW_REQUEST_SHEET_NAME);
  if (!borrowRequestSheet) {
    borrowRequestSheet = ss.insertSheet(BORROW_REQUEST_SHEET_NAME);
    borrowRequestSheet.appendRow(['請求ID', '設備編號', '設備名稱', '借用人', '借用人Email', '借用日期', '預計歸還', '保管人', '狀態', '建立時間']);
  }
  
  const requestId = Utilities.getUuid();
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  
  borrowRequestSheet.appendRow([
    requestId,
    fixNo,
    deviceName,
    borrower,
    borrowerEmail,
    dtBorrow,
    dtDue,
    keeper,
    'pending',
    timestamp
  ]);
  
  // 發送審核郵件給 Keeper
  if (EMAIL_CONFIG.enabled && keeper) {
    sendBorrowApprovalEmail(keeper, fixNo, deviceName, borrower, borrowerEmail, dtBorrow, dtDue, requestId);
  }
  
  Logger.log(`借用請求已建立：${requestId}，等待 Keeper ${keeper} 審核`);
  return successResponse({
    message: '借用申請已送出，等待 Keeper 審核',
    request_id: requestId,
    fix_no: fixNo,
    borrower: borrower,
    keeper: keeper
  });
}

/**
 * 核准借用請求
 */
function approveBorrow(data, e) {
  // 除錯
  Logger.log('=== approveBorrow 開始 ===');
  Logger.log('data:', JSON.stringify(data));
  Logger.log('data type:', typeof data);
  Logger.log('e:', typeof e !== 'undefined' ? 'defined' : 'undefined');
  if (typeof e !== 'undefined' && e.parameter) {
    Logger.log('e.parameter:', JSON.stringify(e.parameter));
  }
  
  // 從 data 或 e 取得 request_id
  let requestId = null;
  if (data && data.request_id) {
    requestId = data.request_id;
  } else if (e && e.parameter && e.parameter.request_id) {
    requestId = e.parameter.request_id;
  }
  
  Logger.log('requestId:', requestId);
  Logger.log('requestId type:', typeof requestId);
  
  if (!requestId) {
    return errorResponse('缺少 request_id');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 查找借用請求
  let borrowRequestSheet = ss.getSheetByName(BORROW_REQUEST_SHEET_NAME);
  if (!borrowRequestSheet) {
    Logger.log('找不到借用申請工作表:', BORROW_REQUEST_SHEET_NAME);
    return errorResponse('找不到借用申請工作表');
  }
  
  const pendingData = borrowRequestSheet.getDataRange().getValues();
  Logger.log('借用申請工作表共有 ' + pendingData.length + ' 列');
  Logger.log('標題列:', JSON.stringify(pendingData[0]));
  
  let foundRow = -1;
  let requestData = null;
  
  for (let i = 1; i < pendingData.length; i++) {
    const rowRequestId = pendingData[i][0];
    Logger.log(`比對第 ${i+1} 列: "${rowRequestId}" (type: ${typeof rowRequestId}) vs "${requestId}" (type: ${typeof requestId})`);
    
    // 使用寬鬆比對（轉為字串）
    if (rowRequestId && rowRequestId.toString() === requestId.toString()) {
      foundRow = i + 1;
      requestData = {
        fix_no: pendingData[i][1],
        device_name: pendingData[i][2],
        borrower: pendingData[i][3],
        borrower_email: pendingData[i][4],
        dt_borrow: pendingData[i][5],
        dt_due: pendingData[i][6],
        keeper: pendingData[i][7]
      };
      Logger.log('找到匹配的請求，資料:', JSON.stringify(requestData));
      break;
    }
  }
  
  if (foundRow === -1 || !requestData) {
    Logger.log('找不到該借用請求或已處理，requestId:', requestId);
    return errorResponse('找不到該借用請求或已處理');
  }
  
  // 更新借用請求狀態為 approved
  borrowRequestSheet.getRange(foundRow, 9).setValue('approved');
  Logger.log('已更新借用請求狀態為 approved');
  
  // 在設備工作表中更新為借用狀態
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  const borrowerCol = COLS.borrower;
  const dtBorrowCol = COLS.dt_borrow;
  const dtDueCol = COLS.dt_due;
  const dtReturnCol = COLS.dt_return;
  const keeperCol = COLS.keeper;
  const deviceNameCol = COLS.device_name;
  
  // 在兩個工作表中查找設備
  Logger.log(`正在查找設備: fix_no='${requestData.fix_no}', SHEET_NAME='${SHEET_NAME}'`);
  Logger.log(`requestData: ${JSON.stringify(requestData)}`);
  let sheet = ss.getSheetByName(SHEET_NAME);
  let equipmentFoundRow = -1;
  let targetSheet = null;
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    Logger.log(`工作表1有 ${lastRow} 行`);
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      Logger.log(`檢查第 ${i} 行: fix_no="${rowFixNo}" vs "${requestData.fix_no}"`);
      if (rowFixNo && rowFixNo.toString().trim() === requestData.fix_no.toString().trim()) {
        equipmentFoundRow = i;
        targetSheet = sheet;
        Logger.log(`在工作表1找到設備 ${requestData.fix_no} 在第 ${i} 行`);
        break;
      }
    }
  } else {
    Logger.log('工作表1不存在');
  }
  
  if (equipmentFoundRow === -1) {
    Logger.log('工作表1找不到設備，嘗試搜尋網站新增設備工作表');
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      Logger.log(`網站新增設備工作表有 ${lastRow} 行`);
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        Logger.log(`檢查第 ${i} 行: fix_no="${rowFixNo}"`);
        if (rowFixNo && rowFixNo.toString().trim() === requestData.fix_no.toString().trim()) {
          equipmentFoundRow = i;
          targetSheet = sheet;
          Logger.log(`在網站新增設備工作表找到設備 ${requestData.fix_no} 在第 ${i} 行`);
          break;
        }
      }
    } else {
      Logger.log('網站新增設備工作表不存在');
    }
  }
  
  // 更新設備為借用狀態
  Logger.log(`equipmentFoundRow=${equipmentFoundRow}, targetSheet=${targetSheet ? targetSheet.getName() : 'null'}`);
  if (equipmentFoundRow !== -1 && targetSheet) {
    // 借用日期以 Keeper 按下同意的日期為準（今天）
    const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
    Logger.log(`借用日期設定為今天: ${today}`);
    
    targetSheet.getRange(equipmentFoundRow, statusCol + 1).setValue('borrowed');
    Logger.log(`已設定 status='borrowed'`);
    targetSheet.getRange(equipmentFoundRow, borrowerCol + 1).setValue(requestData.borrower);
    Logger.log(`已設定 borrower='${requestData.borrower}'`);
    targetSheet.getRange(equipmentFoundRow, dtBorrowCol + 1).setValue(today);  // 使用今天的日期
    Logger.log(`已設定 dt_borrow='${today}'`);
    targetSheet.getRange(equipmentFoundRow, dtDueCol + 1).setValue(requestData.dt_due);
    targetSheet.getRange(equipmentFoundRow, dtReturnCol + 1).setValue('');
    targetSheet.getRange(equipmentFoundRow, COLS.return_confirmed + 1).setValue(false);
    
    const keeper = targetSheet.getRange(equipmentFoundRow, keeperCol + 1).getValue();
    const deviceName = targetSheet.getRange(equipmentFoundRow, deviceNameCol + 1).getValue();
    
    // 記錄歷史（借用日期以今天為準）
    logHistory('borrow', requestData.fix_no, deviceName, requestData.borrower, keeper, today, requestData.dt_due, '');
    
    Logger.log(`設備 ${requestData.fix_no} 已更新為借用狀態，借用人：${requestData.borrower}，借用日期：${today}`);
  } else {
    Logger.log(`找不到設備 ${requestData.fix_no} 來更新狀態！`);
  }
  
  // 發送核准通知給借用人
  if (requestData.borrower_email) {
    sendBorrowResultEmail(requestData.borrower_email, requestData.fix_no, requestData.device_name, requestData.keeper, true);
  }
  
  return successResponse({
    message: '借用已核准',
    fix_no: requestData.fix_no,
    borrower: requestData.borrower
  });
}

/**
 * 拒絕借用請求
 */
function rejectBorrow(data, e) {
  Logger.log('=== rejectBorrow 開始 ===');
  
  // 從 data 或 e 取得 request_id
  let requestId = null;
  if (data && data.request_id) {
    requestId = data.request_id;
  } else if (e && e.parameter && e.parameter.request_id) {
    requestId = e.parameter.request_id;
  }
  
  Logger.log('requestId:', requestId);
  Logger.log('requestId type:', typeof requestId);
  
  if (!requestId) {
    return errorResponse('缺少 request_id');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 查找借用請求
  let borrowRequestSheet = ss.getSheetByName(BORROW_REQUEST_SHEET_NAME);
  if (!borrowRequestSheet) {
    Logger.log('找不到借用申請工作表:', BORROW_REQUEST_SHEET_NAME);
    return errorResponse('找不到借用申請工作表');
  }
  
  const pendingData = borrowRequestSheet.getDataRange().getValues();
  Logger.log('借用申請工作表共有 ' + pendingData.length + ' 列');
  Logger.log('標題列:', JSON.stringify(pendingData[0]));
  
  let foundRow = -1;
  let requestData = null;
  
  for (let i = 1; i < pendingData.length; i++) {
    const rowRequestId = pendingData[i][0];
    Logger.log(`比對第 ${i+1} 列: "${rowRequestId}" (type: ${typeof rowRequestId}) vs "${requestId}" (type: ${typeof requestId})`);
    
    // 使用寬鬆比對（轉為字串）
    if (rowRequestId && rowRequestId.toString() === requestId.toString()) {
      foundRow = i + 1;
      requestData = {
        fix_no: pendingData[i][1],
        device_name: pendingData[i][2],
        borrower: pendingData[i][3],
        borrower_email: pendingData[i][4],
        keeper: pendingData[i][7]
      };
      Logger.log('找到匹配的請求，資料:', JSON.stringify(requestData));
      break;
    }
  }
  
  if (foundRow === -1 || !requestData) {
    Logger.log('找不到該借用請求或已處理，requestId:', requestId);
    return errorResponse('找不到該借用請求或已處理');
  }
  
  // 更新借用請求狀態為 rejected
  borrowRequestSheet.getRange(foundRow, 9).setValue('rejected');
  Logger.log('已更新借用請求狀態為 rejected');
  
  // 拒絕時需要將設備狀態改回 available
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  const borrowerCol = COLS.borrower;
  const dtBorrowCol = COLS.dt_borrow;
  const dtDueCol = COLS.dt_due;
  
  // 在兩個工作表中查找並恢復設備狀態
  Logger.log(`正在查找設備: fix_no='${requestData.fix_no}'`);
  let sheet = ss.getSheetByName(SHEET_NAME);
  let equipmentFoundRow = -1;
  let targetSheet = null;
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    Logger.log(`工作表1有 ${lastRow} 行`);
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      Logger.log(`檢查第 ${i} 行: fix_no="${rowFixNo}" vs "${requestData.fix_no}"`);
      if (rowFixNo && rowFixNo.toString().trim() === requestData.fix_no.toString().trim()) {
        equipmentFoundRow = i;
        targetSheet = sheet;
        Logger.log(`在工作表1找到設備 ${requestData.fix_no} 在第 ${i} 行`);
        break;
      }
    }
  } else {
    Logger.log('工作表1不存在');
  }
  
  if (equipmentFoundRow === -1) {
    Logger.log('工作表1找不到設備，嘗試搜尋網站新增設備工作表');
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      Logger.log(`網站新增設備工作表有 ${lastRow} 行`);
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        Logger.log(`檢查第 ${i} 行: fix_no="${rowFixNo}"`);
        if (rowFixNo && rowFixNo.toString().trim() === requestData.fix_no.toString().trim()) {
          equipmentFoundRow = i;
          targetSheet = sheet;
          Logger.log(`在網站新增設備工作表找到設備 ${requestData.fix_no} 在第 ${i} 行`);
          break;
        }
      }
    } else {
      Logger.log('網站新增設備工作表不存在');
    }
  }
  
  // 恢復設備狀態為可借用
  Logger.log(`equipmentFoundRow=${equipmentFoundRow}, targetSheet=${targetSheet ? targetSheet.getName() : 'null'}`);
  if (equipmentFoundRow !== -1 && targetSheet) {
    Logger.log(`開始恢復設備狀態: 第 ${equipmentFoundRow} 行`);
    targetSheet.getRange(equipmentFoundRow, statusCol + 1).setValue('available');
    Logger.log(`已設定 status='available'`);
    targetSheet.getRange(equipmentFoundRow, borrowerCol + 1).setValue('');
    Logger.log(`已清空 borrower`);
    targetSheet.getRange(equipmentFoundRow, dtBorrowCol + 1).setValue('');
    targetSheet.getRange(equipmentFoundRow, dtDueCol + 1).setValue('');
    Logger.log(`已清空 dt_borrow 和 dt_due`);
    Logger.log(`設備 ${requestData.fix_no} 狀態已恢復為可借用`);
  } else {
    Logger.log(`找不到設備 ${requestData.fix_no} 來恢復狀態！`);
  }
  
  // 發送拒絕通知給借用人
  if (requestData.borrower_email) {
    sendBorrowResultEmail(requestData.borrower_email, requestData.fix_no, requestData.device_name, requestData.keeper, false);
    Logger.log(`已發送拒絕通知給 ${requestData.borrower_email}`);
  }
  
  Logger.log('=== rejectBorrow 完成 ===');
  return successResponse({
    message: '借用已拒絕',
    fix_no: requestData.fix_no,
    borrower: requestData.borrower
  });
}

/**
 * 取得借用請求資訊
 */
function getBorrowRequest(data, e) {
  // 從 data 或 e 取得 request_id
  let requestId = null;
  if (data && data.request_id) {
    requestId = data.request_id;
  } else if (e && e.parameter && e.parameter.request_id) {
    requestId = e.parameter.request_id;
  }
  
  Logger.log('=== getBorrowRequest 開始 ===');
  Logger.log('requestId:', requestId);
  Logger.log('requestId type:', typeof requestId);
  
  if (!requestId) {
    return errorResponse('缺少 request_id');
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 查找借用請求
  let borrowRequestSheet = ss.getSheetByName(BORROW_REQUEST_SHEET_NAME);
  if (!borrowRequestSheet) {
    Logger.log('找不到借用申請工作表:', BORROW_REQUEST_SHEET_NAME);
    return errorResponse('找不到借用申請工作表');
  }
  
  const pendingData = borrowRequestSheet.getDataRange().getValues();
  Logger.log('借用申請工作表共有 ' + pendingData.length + ' 列');
  Logger.log('標題列:', JSON.stringify(pendingData[0]));
  
  for (let i = 1; i < pendingData.length; i++) {
    const rowRequestId = pendingData[i][0];
    Logger.log(`比對第 ${i+1} 列: requestId="${rowRequestId}" (type: ${typeof rowRequestId}) vs 輸入="${requestId}" (type: ${typeof requestId})`);
    
    // 使用寬鬆比對（轉為字串）
    if (rowRequestId && rowRequestId.toString() === requestId.toString()) {
      Logger.log('找到匹配的請求！');
      
      // 檢查狀態，如果已經處理過則回傳失效
      const reqStatus = pendingData[i][8];
      if (reqStatus === 'approved') {
        return errorResponse('此連結已失效（借用已核准）');
      } else if (reqStatus === 'rejected') {
        return errorResponse('此連結已失效（借用已拒絕）');
      } else if (reqStatus !== 'pending') {
        return errorResponse('此連結已失效');
      }
      
      // 格式化日期時間顯示（如果有時間則顯示時間）
      const formatDateTimeForDisplay = (val) => {
        if (!val) return '';
        if (val instanceof Date) {
          return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
        }
        return val.toString();
      };
      
      return successResponse({
        fix_no: pendingData[i][1],
        device_name: pendingData[i][2],
        borrower: pendingData[i][3],
        borrower_email: pendingData[i][4],
        dt_borrow: formatDateTimeForDisplay(pendingData[i][5]),
        dt_due: formatDateTimeForDisplay(pendingData[i][6]),
        keeper: pendingData[i][7],
        status: reqStatus
      });
    }
  }
  
  Logger.log('找不到該借用請求，requestId:', requestId);
  return errorResponse('找不到該借用請求');
}

/**
 * 發送借用審核郵件給 Keeper
 */
function sendBorrowApprovalEmail(keeper, fixNo, deviceName, borrower, borrowerEmail, dtBorrow, dtDue, requestId) {
  try {
    const keeperEmail = getKeeperEmail(keeper);
    
    if (!keeperEmail) {
      Logger.log(`找不到 ${keeper} 的電子郵件`);
      return;
    }
    
    const approveUrl = `${EMAIL_CONFIG.web_app_url}confirm-borrow.html?action=approve&request_id=${encodeURIComponent(requestId)}`;
    const rejectUrl = `${EMAIL_CONFIG.web_app_url}confirm-borrow.html?action=reject&request_id=${encodeURIComponent(requestId)}`;
    
    // 格式化日期時間顯示
    const formatDateTime = (dt) => {
      if (!dt) return '未設定';
      if (dt.includes('T')) {
        const [date, time] = dt.split('T');
        return `${date} ${time}`;
      }
      return dt;
    };
    
    const subject = `${EMAIL_CONFIG.subject_prefix} 借用申請需要您的審核`;
    const body = `親愛的 ${keeper} 您好：

有人申請借用您保管的設備，請審核：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
👤 申請人：${borrower}
📧 申請人 Email：${borrowerEmail}
📅 借用日期：${formatDateTime(dtBorrow)}
⏰ 預計歸還：${formatDateTime(dtDue)}

請點擊以下連結進行審核：
${EMAIL_CONFIG.web_app_url}confirm-borrow.html?request_id=${encodeURIComponent(requestId)}

---
MT 部門設備管理系統 自動通知`.trim();
    
    MailApp.sendEmail(keeperEmail, subject, body);
    Logger.log(`已發送借用審核郵件給 ${keeperEmail}`);
  } catch (err) {
    Logger.error('發送借用審核郵件失敗:', err);
  }
}

/**
 * 發送借用審核結果給借用人
 */
function sendBorrowResultEmail(borrowerEmail, fixNo, deviceName, keeper, isApproved) {
  try {
    const subject = isApproved 
      ? `${EMAIL_CONFIG.subject_prefix} 您的借用申請已核准`
      : `${EMAIL_CONFIG.subject_prefix} 您的借用申請未通過`;
    
    const statusText = isApproved ? '✅ 已核准' : '❌ 未通過';
    const messageText = isApproved 
      ? '您可以前往設備系統查看設備借用狀態。' 
      : '如需借用設備，請聯繫保管人或其他管理員。';
    
    const body = `您好：

您的設備借用申請已有審核結果：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
👤 保管人：${keeper}
📋 審核結果：${statusText}

${messageText}

---
MT 部門設備管理系統 自動通知`.trim();
    
    MailApp.sendEmail(borrowerEmail, subject, body);
    Logger.log(`已發送借用結果通知給 ${borrowerEmail}`);
  } catch (err) {
    Logger.error('發送借用結果郵件失敗:', err);
  }
}

/**
 * 歸還設備
 */
function returnEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const fixNo = data.fix_no;
  // 使用前端傳來的歸還時間（含時間），如果沒有則使用當前台北時間
  let dtReturn = data.dt_return;
  if (!dtReturn) {
    // 如果沒有傳時間，使用當前台北時間（含時間）
    dtReturn = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  }
  
  Logger.log(`歸還設備: ${fixNo}, 歸還時間: ${dtReturn}`);
  
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  const dtReturnCol = COLS.dt_return;
  const keeperCol = COLS.keeper;
  const deviceNameCol = COLS.device_name;
  const borrowerCol = COLS.borrower;
  const dtBorrowCol = COLS.dt_borrow;
  const dtDueCol = COLS.dt_due;
  
  // 先在「工作表 1」查找
  let sheet = ss.getSheetByName(SHEET_NAME);
  let foundRow = -1;
  let targetSheet = null;
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
        foundRow = i;
        targetSheet = sheet;
        break;
      }
    }
  }
  
  // 如果找不到，在「網站新增設備」查找
  if (foundRow === -1) {
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
    }
  }
  
  if (foundRow === -1 || !targetSheet) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const currentStatus = (targetSheet.getRange(foundRow, statusCol + 1).getValue() || '').toString().trim().toLowerCase();
  const isBorrowed = currentStatus === 'borrowed' || currentStatus === '借用中' || currentStatus === '已借出' || currentStatus === '使用中';
  
  if (!isBorrowed) {
    return errorResponse(`設備狀態不是借用中（當前狀態：${currentStatus || '空'}）`);
  }
  
  const keeper = targetSheet.getRange(foundRow, keeperCol + 1).getValue();
  const deviceName = targetSheet.getRange(foundRow, deviceNameCol + 1).getValue();
  const borrower = targetSheet.getRange(foundRow, borrowerCol + 1).getValue();
  const dtBorrowVal = targetSheet.getRange(foundRow, dtBorrowCol + 1).getValue();
  const dtDueVal = targetSheet.getRange(foundRow, dtDueCol + 1).getValue();
  
  targetSheet.getRange(foundRow, dtReturnCol + 1).setValue(dtReturn);
  targetSheet.getRange(foundRow, statusCol + 1).setValue('return_pending');
  
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
  
  const fixNoCol = COLS.fix_no;
  const keeperCol = COLS.keeper;
  
  // 先在「工作表 1」查找
  let sheet = ss.getSheetByName(SHEET_NAME);
  let foundRow = -1;
  let targetSheet = null;
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
        foundRow = i;
        targetSheet = sheet;
        break;
      }
    }
  }
  
  // 如果找不到，在「網站新增設備」查找
  if (foundRow === -1) {
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
    }
  }
  
  if (foundRow === -1 || !targetSheet) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const row = targetSheet.getRange(foundRow, 1, 1, 11).getValues()[0];
  
  // 格式化日期時間顯示（含時間）
  const formatDateTimeForDisplay = (val) => {
    if (!val) return '';
    if (val instanceof Date) {
      return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    }
    // 如果已經是字串，檢查是否有時間部分
    const str = val.toString().trim();
    if (str.includes(' ') || str.includes('T')) {
      // 已經有時間部分
      if (str.includes('T')) {
        const [date, time] = str.split('T');
        return `${date} ${time.substring(0, 5)}`;
      }
      return str.substring(0, 16); // yyyy-MM-dd HH:mm
    }
    // 只有日期，補上 00:00
    return str + ' 00:00';
  };
  
  return successResponse({
    fix_type: row[COLS.fix_type] || '',
    fix_no: row[COLS.fix_no] || '',
    device_name: row[COLS.device_name] || '',
    qty_asset: row[COLS.qty_asset] || '1',
    keeper: row[COLS.keeper] || '',
    status: row[COLS.status] || 'available',
    borrower: row[COLS.borrower] || '',
    dt_borrow: formatDateTimeForDisplay(row[COLS.dt_borrow]),
    dt_due: formatDateTimeForDisplay(row[COLS.dt_due]),
    dt_return: formatDateTimeForDisplay(row[COLS.dt_return]),
    return_confirmed: row[COLS.return_confirmed] || false
  });
}

/**
 * 確認歸還（Keeper 點擊確認連結）
 */
function confirmReturn(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const fixNo = data.fix_no;
  
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  const borrowerCol = COLS.borrower;
  const dtBorrowCol = COLS.dt_borrow;
  const dtDueCol = COLS.dt_due;
  const dtReturnCol = COLS.dt_return;
  const returnConfirmedCol = COLS.return_confirmed;
  const keeperCol = COLS.keeper;
  const deviceNameCol = COLS.device_name;
  
  // 先在「工作表 1」查找
  let sheet = ss.getSheetByName(SHEET_NAME);
  let foundRow = -1;
  let targetSheet = null;
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
        foundRow = i;
        targetSheet = sheet;
        break;
      }
    }
  }
  
  // 如果找不到，在「網站新增設備」查找
  if (foundRow === -1) {
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
    }
  }
  
  if (foundRow === -1 || !targetSheet) {
    return errorResponse(`找不到設備編號：${fixNo}`);
  }
  
  const keeperName = targetSheet.getRange(foundRow, keeperCol + 1).getValue();
  const deviceName = targetSheet.getRange(foundRow, deviceNameCol + 1).getValue();
  const borrower = targetSheet.getRange(foundRow, borrowerCol + 1).getValue();
  const dtBorrowVal = targetSheet.getRange(foundRow, dtBorrowCol + 1).getValue();
  const dtDueVal = targetSheet.getRange(foundRow, dtDueCol + 1).getValue();
  const dtReturnVal = targetSheet.getRange(foundRow, dtReturnCol + 1).getValue();
  
  Logger.log(`確認歸還：${fixNo}，保管人：${keeperName}`);
  
  targetSheet.getRange(foundRow, statusCol + 1).setValue('available');
  targetSheet.getRange(foundRow, returnConfirmedCol + 1).setValue(true);
  targetSheet.getRange(foundRow, borrowerCol + 1).setValue('');
  targetSheet.getRange(foundRow, dtBorrowCol + 1).setValue('');
  targetSheet.getRange(foundRow, dtDueCol + 1).setValue('');
  
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
    
    // 格式化日期時間顯示
    const formatDateTime = (dt) => {
      if (!dt) return '未設定';
      if (dt.includes('T')) {
        const [date, time] = dt.split('T');
        return `${date} ${time}`;
      }
      return dt;
    };
    
    const subject = `${EMAIL_CONFIG.subject_prefix} ${EMAIL_CONFIG.borrow_subject}`;
    const body = `親愛的 ${keeper} 您好：

有人借用了您保管的設備，詳情如下：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
👤 借用人：${borrower}
📅 借用日期：${formatDateTime(dtBorrow)}
⏰ 預計歸還：${formatDateTime(dtDue)}

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

✅ 請點擊以下連結確認歸還：
${EMAIL_CONFIG.web_app_url}/confirm.html?token=${encodeURIComponent(token)}

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
 * 根據姓名從 Keeper 聯絡資訊工作表查找 Email
 */
function getEmailByName(data) {
  try {
    const name = data.name;
    if (!name) {
      return errorResponse('缺少 name 參數');
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const keeperSheet = ss.getSheetByName(KEEPER_SHEET_NAME);
    
    if (!keeperSheet) {
      return errorResponse('找不到 Keeper 聯絡資訊工作表');
    }
    
    const keeperData = keeperSheet.getDataRange().getValues();
    
    // 從第 2 行開始搜尋（跳過標題列）
    for (let i = 1; i < keeperData.length; i++) {
      const rowName = keeperData[i][0];
      const rowEmail = keeperData[i][1];
      
      if (rowName && rowName.toString().trim() === name.toString().trim()) {
        return successResponse({
          name: rowName,
          email: rowEmail || ''
        });
      }
    }
    
    return successResponse({
      name: name,
      email: ''
    });
    
  } catch (err) {
    return errorResponse(err.message);
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
  // 如果 data 是陣列，回傳 { success: true, data: [...] }
  // 否則回傳 { success: true, ...data }
  const output = Array.isArray(data) 
    ? { success: true, data: data }
    : { success: true, ...data };
  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 輔助函式：錯誤回應
 */
function errorResponse(message) {
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    error: message,
    timestamp: new Date().toISOString()
  }))
    .setMimeType(ContentService.MimeType.JSON);
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
  // 防護：確保 data 存在
  if (!data) {
    return errorResponse('參數錯誤：資料為空');
  }
  
  const email = data.email || '';
  const password = data.password || '';
  
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
        Logger.log('收到的密碼："' + password + '"，長度=' + password.length);
        Logger.log('資料庫密碼："' + rowPassword + '"，長度=' + rowPassword.length);
        Logger.log('兩者相等？' + (rowPassword === password));
        
        // 找到匹配的電子郵件，檢查密碼
        if (!rowPassword) {
          Logger.log('密碼為空，需要設定');
          // 密碼為空，回傳需要設定密碼的標記
          return successResponse({
            name: rowName,
            email: rowEmail,
            role: 'admin',
            needSetupPassword: true,
            message: '首次登入，請設定密碼'
          });
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

/**
 * 設定密碼（首次登入）
 */
function setupPassword(data) {
  const email = data.email || '';
  const newPassword = data.newPassword || '';
  
  if (!email || !newPassword) {
    return errorResponse('請提供電子郵件/帳號和新密碼');
  }
  
  if (newPassword.length < 4) {
    return errorResponse('密碼長度至少 4 個字元');
  }
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const keeperSheet = ss.getSheetByName(KEEPER_SHEET_NAME);
    
    if (!keeperSheet) {
      return errorResponse('找不到 Keeper 聯絡資訊工作表');
    }
    
    const sheetData = keeperSheet.getDataRange().getValues();
    
    // 從第 2 列開始（跳過標題）
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const rowEmail = (row[1] || '').toString().trim();      // B 欄 - 電子郵件
      const rowAccount = (row[2] || '').toString().trim();   // C 欄 - 帳號
      const rowName = (row[0] || '').toString().trim();      // A 欄 - 姓名
      
      // 檢查電子郵件或帳號是否匹配
      if (rowEmail === email || rowAccount === email) {
        // 找到匹配，更新密碼（D 欄）
        keeperSheet.getRange(i + 1, 4).setValue(newPassword);  // 第 4 欄 = D 欄
        
        Logger.log('已為 ' + rowName + ' 設定密碼');
        
        return successResponse({
          name: rowName,
          email: rowEmail,
          message: '密碼設定成功'
        });
      }
    }
    
    return errorResponse('找不到此管理員帳號');
    
  } catch (err) {
    Logger.error('設定密碼失敗:', err);
    return errorResponse('設定密碼失敗：' + err.message);
  }
}

/**
 * 更新設備（管理員只能修改自己的設備）
 */
function updateEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const fixNo = data.fix_no;
  
  if (!fixNo) {
    return errorResponse('缺少設備編號');
  }
  
  const fixNoCol = COLS.fix_no;
  const deviceNameCol = COLS.device_name;
  const fixTypeCol = COLS.fix_type;
  const qtyAssetCol = COLS.qty_asset;
  const keeperCol = COLS.keeper;
  
  // 先在「工作表 1」查找
  let sheet = ss.getSheetByName(SHEET_NAME);
  let foundRow = -1;
  let targetSheet = null;
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
        foundRow = i;
        targetSheet = sheet;
        break;
      }
    }
  }
  
  // 如果找不到，在「網站新增設備」查找
  if (foundRow === -1) {
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
    }
  }
  
  if (foundRow === -1 || !targetSheet) {
    return errorResponse('找不到設備編號：' + fixNo);
  }
  
  // 更新資料
  if (data.device_name) {
    targetSheet.getRange(foundRow, deviceNameCol + 1).setValue(data.device_name);
  }
  if (data.fix_type) {
    targetSheet.getRange(foundRow, fixTypeCol + 1).setValue(data.fix_type);
  }
  if (data.qty_asset) {
    targetSheet.getRange(foundRow, qtyAssetCol + 1).setValue(data.qty_asset);
  }
  
  Logger.log('更新設備：' + fixNo + '，更新內容：' + JSON.stringify(data));
  
  return successResponse({
    success: true,
    message: '設備已更新',
    fix_no: fixNo
  });
}

/**
 * 刪除設備（管理員只能刪除自己的設備）
 */
function deleteEquipment(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const fixNo = data.fix_no;
  
  if (!fixNo) {
    return errorResponse('缺少設備編號');
  }
  
  const fixNoCol = COLS.fix_no;
  const statusCol = COLS.status;
  
  // 先在「工作表 1」查找
  let sheet = ss.getSheetByName(SHEET_NAME);
  let foundRow = -1;
  let targetSheet = null;
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
        foundRow = i;
        targetSheet = sheet;
        break;
      }
    }
  }
  
  // 如果找不到，在「網站新增設備」查找
  if (foundRow === -1) {
    sheet = ss.getSheetByName(SHEET_NAME_WEB);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
    }
  }
  
  if (foundRow === -1 || !targetSheet) {
    return errorResponse('找不到設備編號：' + fixNo);
  }
  
  // 檢查設備是否被借出
  const currentStatus = targetSheet.getRange(foundRow, statusCol + 1).getValue();
  const isBorrowed = currentStatus === 'borrowed' || currentStatus === '借用中' || currentStatus === '已借出' || currentStatus === '使用中';
  
  if (isBorrowed) {
    return errorResponse('設備目前被借出，無法刪除');
  }
  
  // 刪除該列
  targetSheet.deleteRow(foundRow);
  
  Logger.log('刪除設備：' + fixNo);
  
  return successResponse({
    success: true,
    message: '設備已刪除',
    fix_no: fixNo
  });
}

// =============================================
// 每日提醒功能
// =============================================

/**
 * 提醒即將到期的設備（預計歸還時間前 1 小時）
 * 建議設定在每小時整點執行（例如 09:00, 10:00...）
 */
function reminderDueSoon() {
  Logger.log('=== 即將到期提醒檢查開始 ===');
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  
  // 強制 now 為整點（分鐘秒設為00）
  now.setMinutes(0, 0, 0);
  
  // 1小時後的整點
  const oneHourLater = new Date(now.getTime() + 60 * 60 * 1000);
  
  Logger.log(`現在時間（整點）: ${Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')}`);
  Logger.log(`1小時後（整點）: ${Utilities.formatDate(oneHourLater, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')}`);
  
  // 讀取兩個工作表
  const sheets = [SHEET_NAME, SHEET_NAME_WEB];
  
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`工作表 ${sheetName} 不存在，跳過`);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`檢查 ${sheetName}，共 ${lastRow} 行`);
    
    for (let i = 2; i <= lastRow; i++) {
      const status = sheet.getRange(i, COLS.status + 1).getValue();
      const dtDue = sheet.getRange(i, COLS.dt_due + 1).getValue();
      
      // 只處理借用中的設備
      if (status !== 'borrowed') {
        continue;
      }
      
      // 跳過沒有預計歸還時間的設備
      if (!dtDue) {
        continue;
      }
      
      // 解析預計歸還時間
      const dueDate = new Date(dtDue);
      if (isNaN(dueDate.getTime())) {
        Logger.log(`無法解析預計歸還時間: ${dtDue}`);
        continue;
      }
      
      // 檢查是否在未來 1 小時內到期
      if (dueDate > now && dueDate <= oneHourLater) {
        const fixNo = sheet.getRange(i, COLS.fix_no + 1).getValue();
        const deviceName = sheet.getRange(i, COLS.device_name + 1).getValue();
        const borrower = sheet.getRange(i, COLS.borrower + 1).getValue();
        
        // 從借用申請工作表查找借用人 Email
        const borrowerEmail = getBorrowerEmailFromRequestSheet(fixNo, borrower);
        
        const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
        Logger.log(`設備 ${fixNo} 將於 1 小時內到期 (${dueTimeStr})，發送提醒給借用人 ${borrower}`);
        sendReminderToBorrower(borrower, borrowerEmail, fixNo, deviceName, dueTimeStr, 'due_soon', now);
      }
    }
  });
  
  Logger.log('=== 即將到期提醒檢查完成 ===');
}

/**
 * 提醒逾期的設備（逾期通知）
 * 建議設定在早上 9:00-10:00 執行
 */
function reminderOverdue() {
  Logger.log('=== 逾期提醒檢查開始 ===');
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  
  // 強制 now 為整點
  now.setMinutes(0, 0, 0);
  
  // 讀取兩個工作表
  const sheets = [SHEET_NAME, SHEET_NAME_WEB];
  
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`工作表 ${sheetName} 不存在，跳過`);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`檢查 ${sheetName}，共 ${lastRow} 行`);
    
    for (let i = 2; i <= lastRow; i++) {
      const status = sheet.getRange(i, COLS.status + 1).getValue();
      const dtDue = sheet.getRange(i, COLS.dt_due + 1).getValue();
      
      // 只處理借用中的設備
      if (status !== 'borrowed') {
        continue;
      }
      
      // 跳過沒有預計歸還時間的設備
      if (!dtDue) {
        continue;
      }
      
      // 解析預計歸還時間
      const dueDate = new Date(dtDue);
      if (isNaN(dueDate.getTime())) {
        Logger.log(`無法解析預計歸還時間: ${dtDue}`);
        continue;
      }
      
      // 只處理已逾期的設備（dueDate 小於現在時間）
      if (dueDate >= now) {
        continue;
      }
      
      const fixNo = sheet.getRange(i, COLS.fix_no + 1).getValue();
      const deviceName = sheet.getRange(i, COLS.device_name + 1).getValue();
      const borrower = sheet.getRange(i, COLS.borrower + 1).getValue();
      const keeper = sheet.getRange(i, COLS.keeper + 1).getValue();
      
      // 從借用申請工作表查找借用人 Email
      const borrowerEmail = getBorrowerEmailFromRequestSheet(fixNo, borrower);
      
      // 計算逾期時間（小時）
      const overdueHours = Math.floor((now - dueDate) / (1000 * 60 * 60));
      const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
      Logger.log(`設備 ${fixNo} 已逾期 ${overdueHours} 小時，發送通知給 Keeper ${keeper} 和借用人 ${borrower}`);
      
      // Keeper 每天收到通知
      sendOverdueNoticeToKeeper(keeper, borrower, fixNo, deviceName, dueTimeStr, now);
      
      // 每天提醒借用人
      sendReminderToBorrower(borrower, borrowerEmail, fixNo, deviceName, dueTimeStr, 'overdue', now);
    }
  });
  
  Logger.log('=== 逾期提醒檢查完成 ===');
}

/**
 * 每日提醒檢查（入口函式）- 只處理普通設備借用
 * 部門儀器提醒請另外設定排程
 */
function dailyReminderCheck() {
  Logger.log('dailyReminderCheck 已棄用，請使用 reminderDueTomorrow 和 reminderOverdue');
  // 同時執行兩個提醒（只處理普通設備，不含部門儀器）
  reminderDueTomorrow();
  reminderOverdue();
}

/**
 * 從借用申請工作表查找借用人 Email
 * @param {string} fixNo - 設備編號
 * @param {string} borrower - 借用人姓名
 * @returns {string|null} 借用人 Email
 */
function getBorrowerEmailFromRequestSheet(fixNo, borrower) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requestSheet = ss.getSheetByName(BORROW_REQUEST_SHEET_NAME);
    
    if (!requestSheet) {
      Logger.log('找不到借用申請工作表');
      return null;
    }
    
    const data = requestSheet.getDataRange().getValues();
    Logger.log(`從借用申請工作表查找 ${borrower} 的 email，共 ${data.length} 列`);
    
    // 從最後一筆開始找（最新的申請）
    for (let i = data.length - 1; i >= 1; i--) {
      const rowFixNo = data[i][1];  // B 欄: 設備編號
      const rowBorrower = data[i][3];  // D 欄: 借用人
      const rowEmail = data[i][4];  // E 欄: 借用人Email
      
      if (rowFixNo && rowFixNo.toString().trim() === fixNo.toString().trim() &&
          rowBorrower && rowBorrower.toString().trim() === borrower.toString().trim()) {
        Logger.log(`找到 ${borrower} 的 email: ${rowEmail}`);
        return rowEmail ? rowEmail.toString().trim() : null;
      }
    }
    
    Logger.log(`在借用申請工作表找不到 ${borrower} 的 email`);
    return null;
  } catch (err) {
    Logger.log(`查找借用人 email 失敗: ${err.message}`);
    return null;
  }
}

/**
 * 發送提醒郵件給借用人
 * @param {string} borrower - 借用人姓名
 * @param {string} borrowerEmail - 借用人 Email
 * @param {string} fixNo - 設備編號
 * @param {string} deviceName - 設備名稱
 * @param {string} dtDue - 預計歸還日期時間
 * @param {string} type - 'due_soon' 或 'overdue'
 * @param {Date} refTime - 參考時間
 */
function sendReminderToBorrower(borrower, borrowerEmail, fixNo, deviceName, dtDue, type, refTime) {
  try {
    if (!borrowerEmail) {
      Logger.log(`找不到 ${borrower} 的電子郵件，無法發送提醒`);
      return;
    }
    
    let subject, body;
    
    if (type === 'due_soon') {
      // 即將到期提醒（1小時內）
      subject = `${EMAIL_CONFIG.subject_prefix} 【提醒】設備將於 1 小時內到期`;
      body = `親愛的 ${borrower} 您好：

提醒您借用的設備即將到期，請準備歸還：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
⏰ 預計歸還：${dtDue}

請在預計歸還時間前歸還設備，謝謝！

---
MT 部門設備管理系統 自動提醒`.trim();
    } else {
      // 逾期提醒
      const dueDate = new Date(dtDue);
      const overdueHours = Math.floor((refTime - dueDate) / (1000 * 60 * 60));
      const overdueDays = Math.floor(overdueHours / 24);
      const remainingHours = overdueHours % 24;
      
      let overdueText = '';
      if (overdueDays > 0) {
        overdueText = `${overdueDays} 天 ${remainingHours} 小時`;
      } else {
        overdueText = `${overdueHours} 小時`;
      }
      
      subject = `${EMAIL_CONFIG.subject_prefix} 【逾期提醒】設備已逾期 ${overdueText}`;
      body = `親愛的 ${borrower} 您好：

您借用的設備已超過預計歸還時間，請盡快歸還：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
⏰ 預計歸還：${dtDue}
⚠️ 逾期時間：${overdueText}

請盡快歸還設備，謝謝！

---
MT 部門設備管理系統 自動提醒`.trim();
    }
    
    MailApp.sendEmail(borrowerEmail, subject, body);
    Logger.log(`已發送提醒郵件給 ${borrower} (${borrowerEmail})`);
  } catch (err) {
    Logger.log(`發送提醒郵件失敗: ${err.message}`);
  }
}

/**
 * 發送逾期通知給 Keeper
 * @param {string} keeper - 保管人姓名
 * @param {string} borrower - 借用人姓名
 * @param {string} fixNo - 設備編號
 * @param {string} deviceName - 設備名稱
 * @param {string} dtDue - 預計歸還日期時間
 * @param {Date} refTime - 參考時間
 */
function sendOverdueNoticeToKeeper(keeper, borrower, fixNo, deviceName, dtDue, refTime) {
  try {
    const keeperEmail = getKeeperEmail(keeper);
    
    if (!keeperEmail) {
      Logger.log(`找不到 ${keeper} 的電子郵件，無法發送通知`);
      return;
    }
    
    const dueDate = new Date(dtDue);
    const overdueHours = Math.floor((refTime - dueDate) / (1000 * 60 * 60));
    const overdueDays = Math.floor(overdueHours / 24);
    const remainingHours = overdueHours % 24;
    
    let overdueText = '';
    if (overdueDays > 0) {
      overdueText = `${overdueDays} 天 ${remainingHours} 小時`;
    } else {
      overdueText = `${overdueHours} 小時`;
    }
    
    const subject = `${EMAIL_CONFIG.subject_prefix} 【通知】借用人逾期未歸還設備`;
    const body = `親愛的 ${keeper} 您好：

借用人 ${borrower} 借用的設備已超過預計歸還時間，尚未歸還：

📦 設備編號：${fixNo}
📝 設備名稱：${deviceName}
👤 借用人：${borrower}
⏰ 預計歸還：${dtDue}
⚠️ 逾期時間：${overdueText}

請聯絡借用人盡快歸還設備。

---
MT 部門設備管理系統 自動通知`.trim();
    
    MailApp.sendEmail(keeperEmail, subject, body);
    Logger.log(`已發送逾期通知給 Keeper ${keeper} (${keeperEmail})`);
  } catch (err) {
    Logger.log(`發送逾期通知失敗: ${err.message}`);
  }
}

/**
 * 部門儀器借用（任何人可用，不需 Keeper 審核）
 * @param {Object} data - 借用資料
 * @param {string} data.device_name - 設備名稱
 * @param {string} data.borrower - 借用人姓名
 * @param {string} data.borrower_email - 借用人郵件
 * @param {string} data.dt_borrow - 借用日期
 * @param {string} data.dt_due - 預計歸還日期
 */
function deptBorrow(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MT部門儀器');
    
    if (!sheet) {
      // 工作表不存在，建立新的
      const newSheet = ss.insertSheet('MT部門儀器');
      // 新增標題列
      newSheet.appendRow([
        'ID', '設備名稱', '借用人', '借用人郵件', '借用日期', '預計歸還', '實際歸還', '狀態'
      ]);
      return deptBorrow(data); // 重新呼叫
    }
    
    const id = Utilities.getUuid().substring(0, 8);
    const timestamp = new Date();
    
    sheet.appendRow([
      id,
      data.device_name,
      data.borrower,
      data.borrower_email,
      data.dt_borrow,
      data.dt_due,
      '', // 實際歸還日期
      '借用中'
    ]);
    
    // 發送確認郵件給借用人
    sendDeptBorrowConfirmation(data.borrower, data.borrower_email, data.device_name, data.dt_borrow, data.dt_due);
    
    return successResponse({
      success: true,
      message: '借用成功',
      id: id
    });
    
  } catch (err) {
    return errorResponse(err.message);
  }
}

/**
 * 部門儀器歸還
 * @param {Object} data - 歸還資料
 * @param {string} data.id - 借用記錄 ID
 */
function deptReturn(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MT部門儀器');
    
    if (!sheet) {
      return errorResponse('找不到 MT部門儀器 工作表');
    }
    
    const allData = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let deviceName = '';
    let borrower = '';
    let borrowerEmail = '';
    let dtBorrow = '';
    let dtDue = '';
    
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.id) {
        rowIndex = i + 1; // 工作表是 1-indexed
        deviceName = allData[i][1];
        borrower = allData[i][2];
        borrowerEmail = allData[i][3];
        dtBorrow = formatDate(allData[i][4]);
        dtDue = formatDate(allData[i][5]);
        break;
      }
    }
    
    if (rowIndex === -1) {
      return errorResponse('找不到該筆借用記錄');
    }
    
    const today = new Date();
    const taipeiTime = new Date(today.getTime() + (8 * 60 * 60 * 1000));
    const dtReturn = taipeiTime.toISOString().split('T')[0];
    
    // 更新狀態
    sheet.getRange(rowIndex, 7).setValue(dtReturn); // G 欄：實際歸還日期
    sheet.getRange(rowIndex, 8).setValue('已歸還'); // H 欄：狀態
    
    // 取得所有管理員郵件
    const adminEmails = getAllAdminEmails();
    
    // 發送歸還通知給所有管理員
    sendDeptReturnNotice(adminEmails, deviceName, borrower, dtBorrow, dtReturn);
    
    return successResponse({
      success: true,
      message: '歸還成功'
    });
    
  } catch (err) {
    return errorResponse(err.message);
  }
}

/**
 * 取得部門儀器借用列表
 */
function getDeptBorrowList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MT部門儀器');
    
    if (!sheet) {
      return successResponse({ success: true, items: [] });
    }
    
    const allData = sheet.getDataRange().getValues();
    const items = [];
    
    for (let i = 1; i < allData.length; i++) {
      items.push({
        id: allData[i][0],
        device_name: allData[i][1],
        borrower: allData[i][2],
        borrower_email: allData[i][3],
        dt_borrow: formatDate(allData[i][4]),
        dt_due: formatDate(allData[i][5]),
        dt_return: formatDate(allData[i][6]),
        status: allData[i][7]
      });
    }
    
    return successResponse({
      success: true,
      items: items
    });
    
  } catch (err) {
    return errorResponse(err.message);
  }
}

/**
 * 取得所有管理員郵件
 */
function getAllAdminEmails() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Keeper 聯絡資訊');
    
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const emails = [];
    
    for (let i = 1; i < data.length; i++) {
      const email = data[i][1];
      if (email && email.includes('@')) {
        emails.push(email);
      }
    }
    
    return emails;
  } catch (err) {
    Logger.log('取得管理員郵件失敗:', err);
    return [];
  }
}

/**
 * 發送部門儀器借用確認郵件
 */
function sendDeptBorrowConfirmation(borrower, borrowerEmail, deviceName, dtBorrow, dtDue) {
  try {
    // 格式化時間顯示（將 T 替換為空格）
    const formatForEmail = (dt) => {
      if (!dt) return '';
      return dt.toString().replace('T', ' ');
    };
    
    const subject = `${EMAIL_CONFIG.subject_prefix} 借用成功確認`;
    const body = `親愛的 ${borrower} 您好：

您已成功借用部門儀器：

📦 設備名稱：${deviceName}
📅 借用日期：${formatForEmail(dtBorrow)}
📅 預計歸還：${formatForEmail(dtDue)}

⚠️ 注意事項：
1. 請於 ${formatForEmail(dtDue)} 前歸還設備
2. 歸還前會收到提醒郵件
3. 如逾期歸還，系統將每天發送提醒通知

感謝您的配合！

---
MT 部門設備管理系統 自動通知`.trim();
    
    MailApp.sendEmail(borrowerEmail, subject, body);
    Logger.log(`已發送借用確認郵件給 ${borrowerEmail}`);
  } catch (err) {
    Logger.log(`發送借用確認郵件失敗: ${err.message}`);
  }
}

/**
 * 發送部門儀器歸還通知給管理員
 */
function sendDeptReturnNotice(adminEmails, deviceName, borrower, dtBorrow, dtReturn) {
  try {
    if (!adminEmails || adminEmails.length === 0) {
      Logger.log('沒有管理員郵件，跳過發送通知');
      return;
    }
    
    const subject = `${EMAIL_CONFIG.subject_prefix} 部門儀器已歸還通知`;
    const body = `各位管理員您好：

部門儀器已歸還，詳情如下：

📦 設備名稱：${deviceName}
👤 借用人：${borrower}
📅 借用日期：${dtBorrow}
📅 歸還日期：${dtReturn}

---
MT 部門設備管理系統 自動通知`.trim();
    
    // 發送給所有管理員
    adminEmails.forEach(email => {
      MailApp.sendEmail(email, subject, body);
    });
    
    Logger.log(`已發送歸還通知給 ${adminEmails.length} 位管理員`);
  } catch (err) {
    Logger.log(`發送歸還通知失敗: ${err.message}`);
  }
}

/**
 * 發送部門儀器逾期提醒（支援小時級別）
 * 找出所有逾期的部門儀器借用，寄信提醒借用人
 */
function sendDeptOverdueReminder() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MT部門儀器');
    
    if (!sheet) {
      Logger.log('找不到 MT部門儀器 工作表');
      return;
    }
    
    const now = new Date();
    const allData = sheet.getDataRange().getValues();
    let sentCount = 0;
    
    for (let i = 1; i < allData.length; i++) {
      const id = allData[i][0];
      const deviceName = allData[i][1];
      const borrower = allData[i][2];
      const borrowerEmail = allData[i][3];
      const dtBorrow = allData[i][4];
      const dtDue = allData[i][5];
      const dtReturn = allData[i][6];
      const status = allData[i][7];
      
      // 只處理狀態是「借用中」且已逾期（歸還時間 < 現在）
      if (status === '借用中' && dtDue && new Date(dtDue) < now && !dtReturn) {
        const dueDate = new Date(dtDue);
        const overdueHours = Math.floor((now - dueDate) / (1000 * 60 * 60));
        const overdueDays = Math.floor(overdueHours / 24);
        const remainingHours = overdueHours % 24;
        
        let overdueText = '';
        if (overdueDays > 0) {
          overdueText = `${overdueDays} 天 ${remainingHours} 小時`;
        } else {
          overdueText = `${overdueHours} 小時`;
        }
        
        const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
        Logger.log(`逾期項目：${deviceName} - ${borrower} (${borrowerEmail}) 已逾期 ${overdueText}`);
        
        if (borrowerEmail && borrowerEmail.includes('@')) {
          const subject = `${EMAIL_CONFIG.subject_prefix} 【逾期提醒】部門儀器已逾期 ${overdueText}`;
          const body = `親愛的 ${borrower} 您好：

您借用的部門儀器已超過預計歸還時間，請盡快歸還：

📦 設備名稱：${deviceName}
⏰ 預計歸還：${dueTimeStr}
⚠️ 逾期時間：${overdueText}

請盡快歸還設備，謝謝！

---
MT 部門設備管理系統 自動提醒`.trim();
          
          MailApp.sendEmail(borrowerEmail, subject, body);
          Logger.log(`已發送逾期提醒給 ${borrower} (${borrowerEmail})`);
          sentCount++;
        }
      }
    }
    
    Logger.log(`逾期提醒發送完畢，共發送 ${sentCount} 封`);
    return sentCount;
    
  } catch (err) {
    Logger.log(`發送部門儀器逾期提醒失敗: ${err.message}`);
    return 0;
  }
}

/**
 * 發送部門儀器歸還前 1 小時的提醒
 * 建議設定為每小時執行
 */
function sendDeptDueSoonReminder() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('MT部門儀器');
    
    if (!sheet) {
      Logger.log('找不到 MT部門儀器 工作表');
      return;
    }
    
    const now = new Date();
    
    // 強制 now 為整點
    now.setMinutes(0, 0, 0);
    
    const oneHourLater = new Date(now.getTime() + 60 * 60 * 1000);
    
    const allData = sheet.getDataRange().getValues();
    let sentCount = 0;
    
    for (let i = 1; i < allData.length; i++) {
      const id = allData[i][0];
      const deviceName = allData[i][1];
      const borrower = allData[i][2];
      const borrowerEmail = allData[i][3];
      const dtBorrow = allData[i][4];
      const dtDue = allData[i][5];
      const dtReturn = allData[i][6];
      const status = allData[i][7];
      
      // 只處理狀態是「借用中」且在 1 小時內到期
      if (status === '借用中' && dtDue && !dtReturn) {
        const dueDate = new Date(dtDue);
        
        // 檢查是否在未來 1 小時內到期
        if (dueDate > now && dueDate <= oneHourLater) {
          const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
          Logger.log(`即將到期：${deviceName} - ${borrower} (歸還時間: ${dueTimeStr})`);
          
          if (borrowerEmail && borrowerEmail.includes('@')) {
            const subject = `${EMAIL_CONFIG.subject_prefix} 【提醒】部門儀器將於 1 小時內到期`;
            const body = `親愛的 ${borrower} 您好：

提醒您借用的部門儀器即將到期，請準備歸還：

📦 設備名稱：${deviceName}
⏰ 預計歸還：${dueTimeStr}

請在預計歸還時間前歸還設備，謝謝！

---
MT 部門設備管理系統 自動提醒`.trim();
            
            MailApp.sendEmail(borrowerEmail, subject, body);
            Logger.log(`已發送即將到期提醒給 ${borrower} (${borrowerEmail})`);
            sentCount++;
          }
        }
      }
    }
    
    Logger.log(`即將到期提醒發送完畢，共發送 ${sentCount} 封`);
    return sentCount;
    
  } catch (err) {
    Logger.log(`發送部門儀器即將到期提醒失敗: ${err.message}`);
    return 0;
  }
}

/**
 * 發送部門儀器歸還前一天的提醒（手動測試用）
 * 已棄用，請使用 sendDeptDueSoonReminder
 */
function sendDeptDueTomorrowReminder() {
  Logger.log('sendDeptDueTomorrowReminder 已棄用，請使用 sendDeptDueSoonReminder');
  sendDeptDueSoonReminder();
}

// =============================================
// 觸發器測試函數
// =============================================

/**
 * 測試提醒功能 - 檢查即將到期設備（不發送郵件）
 * 用法：在 GAS 編輯器中執行 testReminderDueSoon()
 */
function testReminderDueSoon() {
  Logger.log('=== 測試即將到期提醒 ===');
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  now.setMinutes(0, 0, 0);
  
  const oneHourLater = new Date(now.getTime() + 60 * 60 * 1000);
  
  Logger.log(`現在時間（整點）: ${Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')}`);
  Logger.log(`1小時後（整點）: ${Utilities.formatDate(oneHourLater, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')}`);
  
  const sheets = [SHEET_NAME, SHEET_NAME_WEB];
  let foundCount = 0;
  
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ 工作表 ${sheetName} 不存在`);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`\n📋 檢查 ${sheetName}，共 ${lastRow} 行`);
    
    for (let i = 2; i <= lastRow; i++) {
      const status = sheet.getRange(i, COLS.status + 1).getValue();
      const dtDue = sheet.getRange(i, COLS.dt_due + 1).getValue();
      const fixNo = sheet.getRange(i, COLS.fix_no + 1).getValue();
      const deviceName = sheet.getRange(i, COLS.device_name + 1).getValue();
      const borrower = sheet.getRange(i, COLS.borrower + 1).getValue();
      
      if (status !== 'borrowed' || !dtDue) continue;
      
      const dueDate = new Date(dtDue);
      if (isNaN(dueDate.getTime())) {
        Logger.log(`⚠️ 第 ${i} 行: 無法解析時間 "${dtDue}"`);
        continue;
      }
      
      const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
      const isInRange = dueDate > now && dueDate <= oneHourLater;
      
      Logger.log(`  第 ${i} 行: ${fixNo} | ${deviceName} | ${borrower} | ${dueTimeStr} | ${isInRange ? '✅ 符合條件' : '❌ 不符合'}`);
      
      if (isInRange) {
        foundCount++;
        Logger.log(`     📧 將會發送提醒給: ${borrower}`);
      }
    }
  });
  
  Logger.log(`\n🎯 總共找到 ${foundCount} 個設備將在一小時內到期`);
  Logger.log('=== 測試完成 ===');
}

/**
 * 測試提醒功能 - 檢查逾期設備（不發送郵件）
 * 用法：在 GAS 編輯器中執行 testReminderOverdue()
 */
function testReminderOverdue() {
  Logger.log('=== 測試逾期提醒 ===');
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  now.setMinutes(0, 0, 0);
  
  Logger.log(`現在時間（整點）: ${Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')}`);
  
  const sheets = [SHEET_NAME, SHEET_NAME_WEB];
  let foundCount = 0;
  
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`❌ 工作表 ${sheetName} 不存在`);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`\n📋 檢查 ${sheetName}，共 ${lastRow} 行`);
    
    for (let i = 2; i <= lastRow; i++) {
      const status = sheet.getRange(i, COLS.status + 1).getValue();
      const dtDue = sheet.getRange(i, COLS.dt_due + 1).getValue();
      const fixNo = sheet.getRange(i, COLS.fix_no + 1).getValue();
      const deviceName = sheet.getRange(i, COLS.device_name + 1).getValue();
      const borrower = sheet.getRange(i, COLS.borrower + 1).getValue();
      
      if (status !== 'borrowed' || !dtDue) continue;
      
      const dueDate = new Date(dtDue);
      if (isNaN(dueDate.getTime())) continue;
      
      const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
      const isOverdue = dueDate < now;
      
      if (isOverdue) {
        foundCount++;
        const overdueHours = Math.floor((now - dueDate) / (1000 * 60 * 60));
        Logger.log(`  第 ${i} 行: ${fixNo} | ${deviceName} | ${borrower} | ${dueTimeStr} | ⚠️ 逾期 ${overdueHours} 小時`);
      }
    }
  });
  
  Logger.log(`\n🎯 總共找到 ${foundCount} 個逾期設備`);
  Logger.log('=== 測試完成 ===');
}

/**
 * 測試部門儀器提醒 - 檢查即將到期（不發送郵件）
 * 用法：在 GAS 編輯器中執行 testDeptDueSoon()
 */
function testDeptDueSoon() {
  Logger.log('=== 測試部門儀器即將到期 ===');
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  now.setMinutes(0, 0, 0);
  
  const oneHourLater = new Date(now.getTime() + 60 * 60 * 1000);
  
  Logger.log(`現在時間: ${Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm')}`);
  Logger.log(`1小時後: ${Utilities.formatDate(oneHourLater, 'Asia/Taipei', 'yyyy-MM-dd HH:mm')}`);
  
  const sheet = ss.getSheetByName('MT部門儀器');
  if (!sheet) {
    Logger.log('❌ 找不到 MT部門儀器 工作表');
    return;
  }
  
  const allData = sheet.getDataRange().getValues();
  let foundCount = 0;
  
  Logger.log(`\n📋 共 ${allData.length - 1} 筆資料\n`);
  
  for (let i = 1; i < allData.length; i++) {
    const id = allData[i][0];
    const deviceName = allData[i][1];
    const borrower = allData[i][2];
    const borrowerEmail = allData[i][3];
    const dtDue = allData[i][5];
    const status = allData[i][7];
    
    if (status !== '借用中' || !dtDue) continue;
    
    const dueDate = new Date(dtDue);
    if (isNaN(dueDate.getTime())) continue;
    
    const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    const isInRange = dueDate > now && dueDate <= oneHourLater;
    
    Logger.log(`第 ${i} 行: ${deviceName} | ${borrower} | ${dueTimeStr} | ${isInRange ? '✅ 符合' : '❌ 不符合'}`);
    
    if (isInRange) {
      foundCount++;
      Logger.log(`     📧 將會發送提醒給: ${borrower} (${borrowerEmail})`);
    }
  }
  
  Logger.log(`\n🎯 總共找到 ${foundCount} 個部門儀器將在一小時內到期`);
  Logger.log('=== 測試完成 ===');
}

/**
 * 測試部門儀器提醒 - 檢查逾期（不發送郵件）
 * 用法：在 GAS 編輯器中執行 testDeptOverdue()
 */
function testDeptOverdue() {
  Logger.log('=== 測試部門儀器逾期 ===');
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date();
  
  Logger.log(`現在時間: ${Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm')}`);
  
  const sheet = ss.getSheetByName('MT部門儀器');
  if (!sheet) {
    Logger.log('❌ 找不到 MT部門儀器 工作表');
    return;
  }
  
  const allData = sheet.getDataRange().getValues();
  let foundCount = 0;
  
  Logger.log(`\n📋 共 ${allData.length - 1} 筆資料\n`);
  
  for (let i = 1; i < allData.length; i++) {
    const id = allData[i][0];
    const deviceName = allData[i][1];
    const borrower = allData[i][2];
    const borrowerEmail = allData[i][3];
    const dtDue = allData[i][5];
    const dtReturn = allData[i][6];
    const status = allData[i][7];
    
    if (status !== '借用中' || !dtDue || dtReturn) continue;
    
    const dueDate = new Date(dtDue);
    if (isNaN(dueDate.getTime())) continue;
    
    const isOverdue = dueDate < now;
    const dueTimeStr = Utilities.formatDate(dueDate, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    
    if (isOverdue) {
      foundCount++;
      const overdueHours = Math.floor((now - dueDate) / (1000 * 60 * 60));
      Logger.log(`第 ${i} 行: ${deviceName} | ${borrower} | ${dueTimeStr} | ⚠️ 逾期 ${overdueHours} 小時`);
    }
  }
  
  Logger.log(`\n🎯 總共找到 ${foundCount} 個逾期部門儀器`);
  Logger.log('=== 測試完成 ===');
}

/**
 * 一鍵執行所有測試
 * 用法：在 GAS 編輯器中執行 testAllReminders()
 */
function testAllReminders() {
  Logger.log('╔════════════════════════════════════╗');
  Logger.log('║     開始執行所有提醒測試           ║');
  Logger.log('╚════════════════════════════════════╝');
  
  testReminderDueSoon();
  testReminderOverdue();
  testDeptDueSoon();
  testDeptOverdue();
  
  Logger.log('╔════════════════════════════════════╗');
  Logger.log('║     所有測試執行完成！             ║');
  Logger.log('╚════════════════════════════════════╝');
}

// =============================================
// 延後歸還功能
// =============================================

/**
 * 延後預計歸還時間
 * @param {Object} data - 包含 fix_no 和 new_due_date
 * @returns {Object} 成功或失敗的回應
 */
function postponeDueDate(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const fixNo = data.fix_no;
    const newDueDate = data.new_due_date;
    
    Logger.log(`=== 延後歸還時間開始 ===`);
    Logger.log(`設備編號: ${fixNo}, 新時間: ${newDueDate}`);
    
    if (!fixNo || !newDueDate) {
      return errorResponse('缺少設備編號或新的預計歸還時間');
    }
    
    const fixNoCol = COLS.fix_no;
    const statusCol = COLS.status;
    const dtDueCol = COLS.dt_due;
    const keeperCol = COLS.keeper;
    const deviceNameCol = COLS.device_name;
    const borrowerCol = COLS.borrower;
    
    // 先在「工作表 1」查找
    let sheet = ss.getSheetByName(SHEET_NAME);
    let foundRow = -1;
    let targetSheet = null;
    
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
    }
    
    // 如果找不到，在「網站新增設備」查找
    if (foundRow === -1) {
      sheet = ss.getSheetByName(SHEET_NAME_WEB);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        for (let i = 2; i <= lastRow; i++) {
          const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
          if (rowFixNo && rowFixNo.toString().trim() === fixNo) {
            foundRow = i;
            targetSheet = sheet;
            break;
          }
        }
      }
    }
    
    if (foundRow === -1 || !targetSheet) {
      return errorResponse(`找不到設備編號：${fixNo}`);
    }
    
    // 檢查設備狀態是否為借用中
    const currentStatus = targetSheet.getRange(foundRow, statusCol + 1).getValue();
    if (currentStatus !== 'borrowed') {
      return errorResponse(`設備狀態不是借用中（當前狀態：${currentStatus}）`);
    }
    
    // 更新預計歸還時間
    targetSheet.getRange(foundRow, dtDueCol + 1).setValue(newDueDate);
    
    const keeper = targetSheet.getRange(foundRow, keeperCol + 1).getValue();
    const deviceName = targetSheet.getRange(foundRow, deviceNameCol + 1).getValue();
    const borrower = targetSheet.getRange(foundRow, borrowerCol + 1).getValue();
    
    Logger.log(`設備 ${fixNo} 預計歸還時間已更新為: ${newDueDate}`);
    
    // 記錄歷史
    logHistory('postpone', fixNo, deviceName, borrower, keeper, '', newDueDate, '');
    
    return successResponse({
      message: '預計歸還時間已更新',
      fix_no: fixNo,
      new_due_date: newDueDate
    });
    
  } catch (err) {
    Logger.log(`延後歸還時間失敗: ${err.message}`);
    return errorResponse('延後失敗：' + err.message);
  }
}

// =============================================
// 延後歸還申請功能
// =============================================

/**
 * 建立「延後申請」工作表（如果不存在）
 */
function getOrCreatePostponeSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('延後申請');
  if (!sheet) {
    sheet = ss.insertSheet('延後申請');
    sheet.getRange(1, 1, 1, 9).setValues([['request_id', 'fix_no', 'device_name', 'borrower', 'borrower_email', 'current_due', 'new_due', 'dt_request', 'status']]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 發送延後申請（借用人發起）
 */
function requestPostpone(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const { fix_no, new_due_date } = data;
    
    Logger.log(`=== 延後申請開始 ===`);
    Logger.log(`設備編號: ${fix_no}, 新時間: ${new_due_date}`);
    
    if (!fix_no || !new_due_date) {
      return errorResponse('缺少設備編號或新的預計歸還時間');
    }
    
    // 查找設備和借用人資訊
    const fixNoCol = COLS.fix_no;
    const keeperCol = COLS.keeper;
    const deviceNameCol = COLS.device_name;
    const borrowerCol = COLS.borrower;
    const statusCol = COLS.status;
    const dtDueCol = COLS.dt_due;
    
    let sheet = ss.getSheetByName(SHEET_NAME);
    let foundRow = -1;
    let targetSheet = null;
    
    // 在「工作表 1」查找
    if (sheet) {
      const lastRow = sheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fix_no) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
    }
    
    // 在「網站新增設備」查找
    if (foundRow === -1) {
      sheet = ss.getSheetByName(SHEET_NAME_WEB);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        for (let i = 2; i <= lastRow; i++) {
          const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
          if (rowFixNo && rowFixNo.toString().trim() === fix_no) {
            foundRow = i;
            targetSheet = sheet;
            break;
          }
        }
      }
    }
    
    if (foundRow === -1 || !targetSheet) {
      return errorResponse(`找不到設備編號：${fix_no}`);
    }
    
    const keeper = targetSheet.getRange(foundRow, keeperCol + 1).getValue();
    const deviceName = targetSheet.getRange(foundRow, deviceNameCol + 1).getValue();
    const borrower = targetSheet.getRange(foundRow, borrowerCol + 1).getValue();
    const currentDue = targetSheet.getRange(foundRow, dtDueCol + 1).getValue();
    const currentStatus = targetSheet.getRange(foundRow, statusCol + 1).getValue();
    
    if (currentStatus !== 'borrowed') {
      return errorResponse(`設備狀態不是借用中，無法申請延後`);
    }
    
    // 取得借用人 email（優先從「借用申請」工作表找，否則從 Keeper 聯絡資訊）
    let borrowerEmail = '';
    
    // 先從「借用申請」工作表找借用人 email
    const borrowRequestSheet = ss.getSheetByName(BORROW_REQUEST_SHEET_NAME);
    if (borrowRequestSheet) {
      const requestData = borrowRequestSheet.getDataRange().getValues();
      for (let i = 1; i < requestData.length; i++) {
        const reqBorrower = requestData[i][3]; // 借用人欄位
        const reqBorrowerEmail = requestData[i][4]; // 借用人Email欄位
        if (reqBorrower && reqBorrower.toString().trim() === borrower.toString().trim() && reqBorrowerEmail) {
          borrowerEmail = reqBorrowerEmail.toString().trim();
          Logger.log(`從借用申請找到借用人 email: ${borrowerEmail}`);
          break;
        }
      }
    }
    
    // 如果找不到，從 Keeper 聯絡資訊找
    if (!borrowerEmail) {
      borrowerEmail = getKeeperEmail(borrower) || '';
    }
    
    // 產生 request_id
    const requestId = 'PP' + new Date().getTime();
    
    // 寫入延後申請工作表
    const postponeSheet = getOrCreatePostponeSheet();
    const now = new Date();
    const requestTime = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    
    postponeSheet.appendRow([requestId, fix_no, deviceName, borrower, borrowerEmail, currentDue, new_due_date, requestTime, 'pending']);
    
    Logger.log(`延後申請已寫入: ${requestId}`);
    
    // 格式化日期時間（Display purpose）
    const formatDisplayDate = (dateVal) => {
      if (!dateVal) return '未設定';
      if (typeof dateVal === 'string') {
        // 如果是 ISO 字串，轉換 T 為空格
        if (dateVal.includes('T')) {
          const [d, t] = dateVal.split('T');
          return d + ' ' + t.substring(0, 5);
        }
        return dateVal;
      }
      // Date 物件轉換
      const d = new Date(dateVal);
      if (isNaN(d.getTime())) return '未設定';
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const h = String(d.getHours()).padStart(2, '0');
      const min = String(d.getMinutes()).padStart(2, '0');
      return `${y}-${m}-${day} ${h}:${min}`;
    };
    
    // 寄送 email 給 Keeper
    const keeperEmail = getKeeperEmail(keeper) || (keeper && keeper.includes('@') ? keeper : null);
    Logger.log(`延後申請 - Keeper: "${keeper}", Email: "${keeperEmail}"`);
    
    if (keeperEmail) {
      const approveUrl = EMAIL_CONFIG.web_app_url + 'confirm-postpone.html?request_id=' + encodeURIComponent(requestId);
      const subject = EMAIL_CONFIG.subject_prefix + ' 延後歸還申請 - ' + deviceName;
      const currentDueDisplay = formatDisplayDate(currentDue);
      const newDueDisplay = formatDisplayDate(new_due_date);
      const body = `
📋 延後歸還申請通知

設備編號：${fix_no}
設備名稱：${deviceName}
借用人：${borrower}
原預計歸還：${currentDueDisplay}
申請延後至：${newDueDisplay}

請點擊以下連結審核：
${approveUrl}

此郵件由系統自動產生，請勿直接回覆。
      `.trim();
      
      MailApp.sendEmail(keeperEmail, subject, body);
      Logger.log(`✅ 已寄送延後審核郵件給 Keeper: ${keeperEmail}`);
    } else {
      Logger.log(`❌ 找不到 Keeper (${keeper}) 的 email，無法寄送通知`);
      Logger.log(`   請確認「${KEEPER_SHEET_NAME}」工作表中是否有 "${keeper}" 的資料`);
    }
    
    return successResponse({
      message: '延後申請已送出，等待 Keeper 審核',
      request_id: requestId
    });
    
  } catch (err) {
    Logger.log(`延後申請失敗: ${err.message}`);
    return errorResponse('延後申請失敗：' + err.message);
  }
}

/**
 * 取得延後申請資訊（用於審核頁面）
 */
function getPostponeRequest(data, e) {
  try {
    const { request_id } = data;
    
    Logger.log(`取得延後申請: ${request_id}`);
    
    const postponeSheet = getOrCreatePostponeSheet();
    const lastRow = postponeSheet.getLastRow();
    
    for (let i = 2; i <= lastRow; i++) {
      const rowRequestId = postponeSheet.getRange(i, 1).getValue();
      if (rowRequestId && rowRequestId.toString() === request_id) {
        const fixNo = postponeSheet.getRange(i, 2).getValue();
        const deviceName = postponeSheet.getRange(i, 3).getValue();
        const borrower = postponeSheet.getRange(i, 4).getValue();
        const currentDue = postponeSheet.getRange(i, 6).getValue();
        const newDue = postponeSheet.getRange(i, 7).getValue();
        const status = postponeSheet.getRange(i, 9).getValue();
        
        if (status === 'approved') {
          return errorResponse('此連結已失效（申請已核准）');
        } else if (status === 'rejected') {
          return errorResponse('此連結已失效（申請已拒絕）');
        } else if (status === 'pending') {
          // 還在等待審核，這是正確的
          return successResponse({
            fix_no: fixNo,
            device_name: deviceName,
            borrower: borrower,
            current_due: currentDue,
            new_due: newDue,
            status: status
          });
        } else {
          return errorResponse('此連結已失效');
        }
        
        return successResponse({
          fix_no: fixNo,
          device_name: deviceName,
          borrower: borrower,
          current_due: currentDue,
          new_due: newDue,
          status: status
        });
      }
    }
    
    return errorResponse('找不到此延後申請');
    
  } catch (err) {
    Logger.log(`取得延後申請失敗: ${err.message}`);
    return errorResponse('載入失敗：' + err.message);
  }
}

/**
 * 核准延後申請（Keeper 按下同意）
 */
function approvePostpone(data, e) {
  try {
    const { request_id } = data;
    
    Logger.log(`核准延後申請: ${request_id}`);
    
    const postponeSheet = getOrCreatePostponeSheet();
    const lastRow = postponeSheet.getLastRow();
    let foundRow = -1;
    
    for (let i = 2; i <= lastRow; i++) {
      const rowRequestId = postponeSheet.getRange(i, 1).getValue();
      if (rowRequestId && rowRequestId.toString() === request_id) {
        foundRow = i;
        break;
      }
    }
    
    if (foundRow === -1) {
      return errorResponse('找不到此延後申請');
    }
    
    const status = postponeSheet.getRange(foundRow, 9).getValue();
    if (status !== 'pending') {
      return errorResponse('此申請已經處理過了');
    }
    
    const fixNo = postponeSheet.getRange(foundRow, 2).getValue();
    const deviceName = postponeSheet.getRange(foundRow, 3).getValue();
    const borrower = postponeSheet.getRange(foundRow, 4).getValue();
    const borrowerEmail = postponeSheet.getRange(foundRow, 5).getValue();
    const newDue = postponeSheet.getRange(foundRow, 7).getValue();
    
    // 更新申請狀態為 approved
    postponeSheet.getRange(foundRow, 9).setValue('approved');
    
    // 更新設備的 dt_due
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const fixNoCol = COLS.fix_no;
    const dtDueCol = COLS.dt_due;
    const keeperCol = COLS.keeper;
    
    // 查找設備所在的 sheet 並更新
    let sheet = ss.getSheetByName(SHEET_NAME);
    let targetRow = -1;
    let foundInSheet = '';
    
    Logger.log(`在 ${SHEET_NAME} 中搜尋設備 ${fixNo}...`);
    for (let i = 2; i <= sheet.getLastRow(); i++) {
      const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
      Logger.log(`第 ${i} 列: 值="${rowFixNo}" (${typeof rowFixNo}), 比對="${fixNo}"`);
      if (rowFixNo && rowFixNo.toString().trim() === fixNo.toString().trim()) {
        targetRow = i;
        foundInSheet = SHEET_NAME;
        Logger.log(`✅ 在 ${SHEET_NAME} 第 ${i} 列找到`);
        break;
      }
    }
    
    if (targetRow === -1) {
      Logger.log(`在 ${SHEET_NAME} 找不到，搜尋 ${SHEET_NAME_WEB}...`);
      sheet = ss.getSheetByName(SHEET_NAME_WEB);
      for (let i = 2; i <= sheet.getLastRow(); i++) {
        const rowFixNo = sheet.getRange(i, fixNoCol + 1).getValue();
        if (rowFixNo && rowFixNo.toString().trim() === fixNo.toString().trim()) {
          targetRow = i;
          foundInSheet = SHEET_NAME_WEB;
          Logger.log(`✅ 在 ${SHEET_NAME_WEB} 第 ${i} 列找到`);
          break;
        }
      }
    }
    
    if (targetRow !== -1) {
      Logger.log(`更新設備 ${fixNo} 的 dt_due: ${newDue}`);
      sheet.getRange(targetRow, dtDueCol + 1).setValue(newDue);
      const keeper = sheet.getRange(targetRow, keeperCol + 1).getValue();
      logHistory('postpone_approved', fixNo, deviceName, borrower, keeper, '', newDue, '');
      Logger.log(`✅ dt_due 更新成功`);
    } else {
      Logger.log(`❌ 找不到設備 ${fixNo} 的列，無法更新 dt_due`);
    }
    
    // 格式化日期時間
    const formatDisplayDate = (dateVal) => {
      if (!dateVal) return '未設定';
      if (typeof dateVal === 'string') {
        if (dateVal.includes('T')) {
          const [d, t] = dateVal.split('T');
          return d + ' ' + t.substring(0, 5);
        }
        return dateVal;
      }
      const d = new Date(dateVal);
      if (isNaN(d.getTime())) return '未設定';
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const h = String(d.getHours()).padStart(2, '0');
      const min = String(d.getMinutes()).padStart(2, '0');
      return `${y}-${m}-${day} ${h}:${min}`;
    };
    
    // 寄送核准通知給借用人
    if (borrowerEmail) {
      const newDueDisplay = formatDisplayDate(newDue);
      const subject = EMAIL_CONFIG.subject_prefix + ' 延後歸還已核准 - ' + deviceName;
      const body = `
✅ 延後歸還申請已核准

設備編號：${fixNo}
設備名稱：${deviceName}
新預計歸還時間：${newDueDisplay}

請在新的預計歸還時間前歸還設備。

此郵件由系統自動產生，請勿直接回覆。
      `.trim();
      
      MailApp.sendEmail(borrowerEmail, subject, body);
    }
    
    return successResponse({
      message: '已核准延後申請',
      fix_no: fixNo,
      new_due: newDue
    });
    
  } catch (err) {
    Logger.log(`核准延後失敗: ${err.message}`);
    return errorResponse('核准失敗：' + err.message);
  }
}

/**
 * 拒絕延後申請（Keeper 按下不同意）
 */
function rejectPostpone(data, e) {
  try {
    const { request_id } = data;
    
    Logger.log(`拒絕延後申請: ${request_id}`);
    
    const postponeSheet = getOrCreatePostponeSheet();
    const lastRow = postponeSheet.getLastRow();
    let foundRow = -1;
    
    for (let i = 2; i <= lastRow; i++) {
      const rowRequestId = postponeSheet.getRange(i, 1).getValue();
      if (rowRequestId && rowRequestId.toString() === request_id) {
        foundRow = i;
        break;
      }
    }
    
    if (foundRow === -1) {
      return errorResponse('找不到此延後申請');
    }
    
    const status = postponeSheet.getRange(foundRow, 9).getValue();
    if (status !== 'pending') {
      return errorResponse('此申請已經處理過了');
    }
    
    const fixNo = postponeSheet.getRange(foundRow, 2).getValue();
    const deviceName = postponeSheet.getRange(foundRow, 3).getValue();
    const borrower = postponeSheet.getRange(foundRow, 4).getValue();
    const borrowerEmail = postponeSheet.getRange(foundRow, 5).getValue();
    const currentDue = postponeSheet.getRange(foundRow, 6).getValue();
    
    // 更新申請狀態為 rejected
    postponeSheet.getRange(foundRow, 9).setValue('rejected');
    
    // 格式化日期時間
    const formatDisplayDate = (dateVal) => {
      if (!dateVal) return '未設定';
      if (typeof dateVal === 'string') {
        if (dateVal.includes('T')) {
          const [d, t] = dateVal.split('T');
          return d + ' ' + t.substring(0, 5);
        }
        return dateVal;
      }
      const d = new Date(dateVal);
      if (isNaN(d.getTime())) return '未設定';
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const h = String(d.getHours()).padStart(2, '0');
      const min = String(d.getMinutes()).padStart(2, '0');
      return `${y}-${m}-${day} ${h}:${min}`;
    };
    
    // 寄送拒絕通知給借用人
    if (borrowerEmail) {
      const currentDueDisplay = formatDisplayDate(currentDue);
      const subject = EMAIL_CONFIG.subject_prefix + ' 延後歸還已拒絕 - ' + deviceName;
      const body = `
❌ 延後歸還申請已拒絕

設備編號：${fixNo}
設備名稱：${deviceName}
維持原本預計歸還時間：${currentDueDisplay}

請在原本的預計歸還時間前歸還設備。

如有疑問，請聯繫 Keeper。

此郵件由系統自動產生，請勿直接回覆。
      `.trim();
      
      MailApp.sendEmail(borrowerEmail, subject, body);
    }
    
    return successResponse({
      message: '已拒絕延後申請',
      fix_no: fixNo
    });
    
  } catch (err) {
    Logger.log(`拒絕延後失敗: ${err.message}`);
    return errorResponse('拒絕失敗：' + err.message);
  }
}
