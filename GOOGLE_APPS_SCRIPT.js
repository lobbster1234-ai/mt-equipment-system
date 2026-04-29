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
 const actionType = params.actionType || ''; // 修正：使用 actionType 避免與 action='history' 衝突
 
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
 
 // 欄位索引：0=時間戳，1=動作，2=設備編號，3=設備名稱，4=借用人，5=保管人，6=借用日期，7=預計歸還，8=實際歸還/確認日期
 // 直接回傳原始欄位名稱，讓前端統一處理
 return {
 timestamp: row[0] || '',
 action: action,
 fix_no: row[2] || '',
 device_name: row[3] || '',
 borrower: row[4] || '',
 keeper: row[5] || '',
 dt_borrow: formatDate(row[6]),
 dt_due: formatDate(row[7]),
 dt_return: formatDate(row[8]),
 return_confirmed: (action === 'confirm' || action === 'confirmed') ? true : false
 };
 });
 
 return successResponse(result);
}
