// 簡化版 GAS - 移除待審核借用工作表
// 只保留核心功能：借用、歸還、審核

function doGet(e) {
  try {
    const action = e.parameter.action || 'query';
    
    if (action === 'query') {
      return queryEquipment(e.parameter);
    } else if (action === 'borrow') {
      // 借用申請 - 直接設為 borrow_pending
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const fixNo = e.parameter.fix_no;
      const borrower = e.parameter.borrower;
      const dtBorrow = e.parameter.dt_borrow || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
      const dtDue = e.parameter.dt_due || '';
      
      // 查找設備
      let sheet = ss.getSheetByName(SHEET_NAME);
      let foundRow = -1;
      let targetSheet = null;
      
      for (let i = 2; i <= sheet.getLastRow(); i++) {
        if (sheet.getRange(i, COLS.fix_no + 1).getValue() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
      
      if (foundRow === -1) {
        sheet = ss.getSheetByName(SHEET_NAME_WEB);
        for (let i = 2; i <= sheet.getLastRow(); i++) {
          if (sheet.getRange(i, COLS.fix_no + 1).getValue() === fixNo) {
            foundRow = i;
            targetSheet = sheet;
            break;
          }
        }
      }
      
      if (!targetSheet) return errorResponse('找不到設備');
      
      const currentStatus = targetSheet.getRange(foundRow, COLS.status + 1).getValue();
      if (currentStatus === 'borrowed' || currentStatus === 'borrow_pending') {
        return errorResponse('設備已借出或審核中');
      }
      
      const keeper = targetSheet.getRange(foundRow, COLS.keeper + 1).getValue();
      const deviceName = targetSheet.getRange(foundRow, COLS.device_name + 1).getValue();
      
      // 更新狀態為借用審核中
      targetSheet.getRange(foundRow, COLS.status + 1).setValue('borrow_pending');
      targetSheet.getRange(foundRow, COLS.borrower + 1).setValue(borrower);
      targetSheet.getRange(foundRow, COLS.dt_borrow + 1).setValue(dtBorrow);
      targetSheet.getRange(foundRow, COLS.dt_due + 1).setValue(dtDue);
      
      // 記錄歷史
      logHistory('borrow_pending', fixNo, deviceName, borrower, keeper, dtBorrow, dtDue, '');
      
      // 發送郵件給 Keeper
      if (EMAIL_CONFIG.enabled && keeper) {
        const keeperEmail = getKeeperEmail(keeper);
        const subject = `${EMAIL_CONFIG.subject_prefix} 借用申請需要審核`;
        const body = `親愛的 ${keeper} 您好：

${borrower} 申請借用 ${deviceName} (${fixNo})

請點擊以下連結審核：
同意: ${EMAIL_CONFIG.web_app_url}/confirm.html?action=approve&fix_no=${encodeURIComponent(fixNo)}
拒絕: ${EMAIL_CONFIG.web_app_url}/confirm.html?action=reject&fix_no=${encodeURIComponent(fixNo)}`;
        MailApp.sendEmail(keeperEmail, subject, body);
      }
      
      return successResponse({
        message: '借用申請已送出，等待 Keeper 審核',
        fix_no: fixNo,
        borrower: borrower
      });
      
    } else if (action === 'approveBorrow') {
      // Keeper 同意借用
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const fixNo = e.parameter.fix_no;
      
      // 查找設備
      let sheet = ss.getSheetByName(SHEET_NAME);
      let foundRow = -1;
      let targetSheet = null;
      
      for (let i = 2; i <= sheet.getLastRow(); i++) {
        if (sheet.getRange(i, COLS.fix_no + 1).getValue() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
      
      if (foundRow === -1) {
        sheet = ss.getSheetByName(SHEET_NAME_WEB);
        for (let i = 2; i <= sheet.getLastRow(); i++) {
          if (sheet.getRange(i, COLS.fix_no + 1).getValue() === fixNo) {
            foundRow = i;
            targetSheet = sheet;
            break;
          }
        }
      }
      
      if (!targetSheet) return errorResponse('找不到設備');
      
      const borrower = targetSheet.getRange(foundRow, COLS.borrower + 1).getValue();
      const keeper = targetSheet.getRange(foundRow, COLS.keeper + 1).getValue();
      const deviceName = targetSheet.getRange(foundRow, COLS.device_name + 1).getValue();
      const dtBorrow = targetSheet.getRange(foundRow, COLS.dt_borrow + 1).getValue();
      const dtDue = targetSheet.getRange(foundRow, COLS.dt_due + 1).getValue();
      
      // 更新為借用中
      targetSheet.getRange(foundRow, COLS.status + 1).setValue('borrowed');
      
      // 記錄歷史
      logHistory('borrow', fixNo, deviceName, borrower, keeper, dtBorrow, dtDue, '');
      
      return successResponse({
        message: '借用已核准',
        fix_no: fixNo,
        borrower: borrower
      });
      
    } else if (action === 'rejectBorrow') {
      // Keeper 拒絕借用
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const fixNo = e.parameter.fix_no;
      
      // 查找設備
      let sheet = ss.getSheetByName(SHEET_NAME);
      let foundRow = -1;
      let targetSheet = null;
      
      for (let i = 2; i <= sheet.getLastRow(); i++) {
        if (sheet.getRange(i, COLS.fix_no + 1).getValue() === fixNo) {
          foundRow = i;
          targetSheet = sheet;
          break;
        }
      }
      
      if (foundRow === -1) {
        sheet = ss.getSheetByName(SHEET_NAME_WEB);
        for (let i = 2; i <= sheet.getLastRow(); i++) {
          if (sheet.getRange(i, COLS.fix_no + 1).getValue() === fixNo) {
            foundRow = i;
            targetSheet = sheet;
            break;
          }
        }
      }
      
      if (!targetSheet) return errorResponse('找不到設備');
      
      // 恢復為可借用
      targetSheet.getRange(foundRow, COLS.status + 1).setValue('available');
      targetSheet.getRange(foundRow, COLS.borrower + 1).setValue('');
      targetSheet.getRange(foundRow, COLS.dt_borrow + 1).setValue('');
      targetSheet.getRange(foundRow, COLS.dt_due + 1).setValue('');
      
      return successResponse({
        message: '借用已拒絕，設備恢復可借用',
        fix_no: fixNo
      });
    }
    
    return errorResponse('未知的 action: ' + action);
  } catch (err) {
    return errorResponse(err.message);
  }
}

// 保留其他必要的輔助函數...