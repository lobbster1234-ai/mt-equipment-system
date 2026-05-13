// 測試腳本：檢查 borrow_pending 狀態處理
console.log('=== 測試 borrow_pending 狀態 ===');

// 模擬設備資料
const testEquipment = [
  { fix_no: 'E001', device_name: '測試設備1', status: 'available', borrower: '' },
  { fix_no: 'E002', device_name: '測試設備2', status: 'borrow_pending', borrower: '測試用戶' },
  { fix_no: 'E003', device_name: '測試設備3', status: 'borrowed', borrower: '測試用戶2', dt_borrow: '2025-05-13' },
];

// 測試渲染邏輯
testEquipment.forEach(eq => {
  const isAvailable = eq.status === 'available' || eq.status === '可借用' || !eq.status;
  const isBorrowPending = eq.status === 'borrow_pending' || eq.status === '借用審核中';
  const isReturnPending = eq.status === 'return_pending' || eq.status === '歸還認證中';
  const isConfirmed = eq.return_confirmed === true || eq.return_confirmed === 'true' || eq.return_confirmed === 1;
  
  let statusHtml;
  if (isAvailable) {
    statusHtml = '✅ 可借用';
  } else if (isBorrowPending) {
    statusHtml = '⏳ 借用審核中';
  } else if (isReturnPending) {
    statusHtml = '⏳ 歸還認證中';
  } else if (isConfirmed) {
    statusHtml = '✅ 已確認';
  } else {
    statusHtml = '📤 借用中';
  }
  
  let actionButton = '';
  if (isAvailable) {
    actionButton = '[借用按鈕]';
  } else if (isBorrowPending) {
    actionButton = '[⏳ 等待 Keeper 審核]';
  } else if (isReturnPending) {
    actionButton = '[等待 Keeper 確認]';
  } else if (!isConfirmed) {
    actionButton = '[📧 歸還按鈕]';
  } else {
    actionButton = '[已確認]';
  }
  
  console.log(`\n設備 ${eq.fix_no}:`);
  console.log(`  狀態: ${eq.status}`);
  console.log(`  顯示: ${statusHtml}`);
  console.log(`  按鈕: ${actionButton}`);
});

console.log('\n=== 測試完成 ===');
