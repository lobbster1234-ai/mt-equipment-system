// =============================================
// MT 設備系統 - Google Apps Script 前端
// =============================================

const GAS_URL = 'https://script.google.com/macros/s/AKfycbxeI5xC33a6Ry634g6kwBPK9feElH_tTPtQYeWcH4ReiEiiq5I9yIetv8ugAFDgJkHh1A/exec';

/**
 * 格式化日期時間為 yyyy-MM-dd HH:mm:ss（處理各種輸入格式）
 */
function formatDateTime(value) {
  if (!value) return '';
  
  // 如果已經是 yyyy-MM-dd HH:mm:ss 格式，直接返回
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?$/.test(value.trim())) {
    return value.trim();
  }
  
  // 處理 JavaScript Date 物件字串格式："Thu Apr 30 2026 00:00:00 GMT+0800 (台北標準時間)"
  if (typeof value === 'string' && value.includes('GMT')) {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
  }
  
  // 嘗試解析為 Date
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return String(value).trim();
  }
  
  // 格式化為 yyyy-MM-dd HH:mm:ss（台北時間 GMT+8）
  const taipeiOffset = 8 * 60;
  const localOffset = date.getTimezoneOffset();
  const offsetDiff = taipeiOffset + localOffset;
  const taipeiDate = new Date(date.getTime() + offsetDiff * 60000);
  
  const year = taipeiDate.getFullYear();
  const month = String(taipeiDate.getMonth() + 1).padStart(2, '0');
  const day = String(taipeiDate.getDate()).padStart(2, '0');
  const hours = String(taipeiDate.getHours()).padStart(2, '0');
  const minutes = String(taipeiDate.getMinutes()).padStart(2, '0');
  const seconds = String(taipeiDate.getSeconds()).padStart(2, '0');
  
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}
