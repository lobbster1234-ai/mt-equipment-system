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
  
  // 如果是 00:00:00 就只顯示日期，否則顯示日期+時間
  if (hours === '00' && minutes === '00' && seconds === '00') {
    return `${year}-${month}-${day}`;
  }
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// 重要：GAS 部署設定必須為：
// 1. 執行身分：我 (stella_fan)
// 2. 誰有權存取：任何人 (Anyone, even anonymous)
// 3. 部署後網址要用 /exec 結尾，不是 /dev

// 測試 GAS 是否正常：直接在瀏覽器開啟 GAS_URL，應該要看到 JSON 回應而不是 401/403

// =============================================
// 查詢功能
// =============================================

// 搜尋設備
async function searchEquipment() {
  const keyword = document.getElementById('search-keyword')?.value.trim() || '';
  const department = document.getElementById('search-department')?.value || '';
  const status = document.getElementById('search-status')?.value || '';

  const params = new URLSearchParams({ action: 'query' });
  if (keyword) params.append('keyword', keyword);
  if (department) params.append('dept_id', department);
  if (status) params.append('status', status);

  const listEl = document.getElementById('equipment-list');
  if (listEl) {
    listEl.innerHTML = '<p style="text-align:center;color:#666;padding:40px;">🔄 載入中...</p>';
  }

  try {
    // 使用 GET 請求 + redirect=follow 避免 CORS 問題
    // GAS 會 302 重定向，follow 會自動跟隨
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'query');
    if (keyword) url.searchParams.append('keyword', keyword);
    if (department) url.searchParams.append('dept_id', department);
    if (status) url.searchParams.append('status', status);

    console.log('GAS 請求網址:', url.toString());

    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });

    console.log('GAS 回應狀態:', res.status, res.ok);

  if (!res.ok) {
      const errorText = await res.text();
      console.error('GAS 錯誤回應:', errorText);
      let message = `HTTP ${res.status}: ${res.statusText}.`;
      // 由於您提到系統在 GitHub 部署，請確認 GAS_URL 是否已更新為您新的 API 端點。
      message += `如果 API 網址錯誤，請檢查 URL。原始錯誤: ${res.statusText}`;
      throw new Error(message);
  }

    const data = await res.json();
    console.log('GAS 回應資料:', data);

    if (data.error) {
      throw new Error(data.error);
    }

    // 處理陣列或物件格式
    const equipment = Array.isArray(data) ? data : (data.data || data.result || data.items || []);
    console.log('設備資料:', equipment.slice(0, 5)); // 顯示前 5 筆
    renderEquipment(equipment);
  } catch (err) {
    console.error('查詢失敗:', err);
    if (listEl) {
      // 提供 CORS 錯誤提示
      let msg = err.message;
      if (msg.includes('Failed to fetch') || msg.includes('NetworkError')) {
        msg = 'CORS 錯誤 - 請確認 GAS 部署設定為「誰有權存取：任何人」';
      }
      listEl.innerHTML = `<p style="text-align:center;color:#c00;padding:40px;">❌ 查詢失敗：${msg}</p>`;
    }
  }
}

// 渲染設備列表
function renderEquipment(equipment) {
  const list = document.getElementById('equipment-list');
  if (!list) return;

  if (!equipment || equipment.length === 0) {
    list.innerHTML = '<p style="text-align:center;color:#666;padding:40px;">目前沒有設備</p>';
    return;
  }

  // 按保管人分組
  const grouped = {};
  equipment.forEach(eq => {
    const fixNo = eq.fix_no || '';
    const deviceName = eq.device_name || '';
    const keeper = eq.keeper || '未指定';

    if (!grouped[keeper]) grouped[keeper] = [];
    grouped[keeper].push({ ...eq, fix_no: fixNo, device_name: deviceName, keeper: keeper });
  });

  // 產生 HTML
  let html = '';
  Object.keys(grouped).sort().forEach((keeper, index) => {
    const items = grouped[keeper];
    const isExpanded = index === 0; // 第一個預設展開
    
    html += `
      <div class="keeper-group">
        <div class="keeper-header" onclick="toggleKeeperGroup(this)" style="cursor:pointer;user-select:none;">
          <span class="keeper-arrow" style="display:inline-block;width:12px;margin-right:8px;transition:transform 0.2s;${isExpanded ? 'transform:rotate(90deg)' : ''}">▶</span>
          <span>${getAvatarHtml(keeper, 55)} ${keeper} (${items.length}項)</span>
        </div>
        <div class="keeper-table-wrapper" style="${isExpanded ? 'display:block;' : 'display:none;'}">
          <table class="equipment-table">
            <thead>
              <tr>
                <th>設備類型</th>
                <th>設備編號</th>
                <th>設備名稱</th>
                <th>數量</th>
                <th>狀態</th>
              </tr>
            </thead>
            <tbody>
              ${items.map(eq => {
                const isAvailable = eq.status === 'available' || eq.status === '可借用' || !eq.status;
                const isReturnPending = eq.status === 'return_pending' || eq.status === '歸還認證中';
                const isConfirmed = eq.return_confirmed === true || eq.return_confirmed === 'true' || eq.return_confirmed === 1;
                
                let statusHtml;
                if (isAvailable) {
                  statusHtml = '<span style="color:#0a0;">✅ 可借用</span>';
                } else if (isReturnPending) {
                  statusHtml = '<span style="color:#17a2b8;">⏳ 歸還認證中</span>';
                } else if (isConfirmed) {
                  statusHtml = '<span style="color:#999;">✅ 已確認</span>';
                } else {
                  statusHtml = '<span style="color:#c00;">📤 借用中</span>';
                }
                
                // 借用/歸還按鈕
                let actionButton = '';
                // URL 編碼設備名稱，避免特殊字元破壞 HTML/JavaScript 語法
                const encodedDeviceName = encodeURIComponent(eq.device_name || '');
                if (isAvailable) {
                  actionButton = `<button class="btn-borrow-sm" onclick="openBorrowModal('${eq.fix_no}', decodeURIComponent('${encodedDeviceName}'), '${eq.keeper}')">借用</button>`;
                } else if (isReturnPending) {
                  // 歸還認證中，不顯示按鈕
                  actionButton = '<span style="color:#17a2b8;font-size:0.85em;">等待 Keeper 確認</span>';
                } else if (!isConfirmed) {
                  // 借用中，未歸還
                  const hasReturnDate = eq.dt_return && eq.dt_return.toString().trim() !== '';
                  if (hasReturnDate) {
                    actionButton = '<span style="color:#17a2b8;font-size:0.85em;">⏳ 待確認</span>';
                  } else {
                    actionButton = `<button class="btn-return-sm" onclick="openReturnModal('${eq.fix_no}', decodeURIComponent('${encodedDeviceName}'), '${eq.borrower}')">📧 歸還</button>`;
                  }
                } else {
                  actionButton = '<span style="color:#999;font-size:0.85em;">已確認</span>';
                }
                
                return `
                  <tr>
                    <td>${eq.fix_type || ''}</td>
                    <td>${eq.fix_no || ''}</td>
                    <td>${eq.device_name || ''}</td>
                    <td>${eq.qty_asset || '1'}</td>
                    <td>
                      ${statusHtml}
                      <div style="margin-top:5px;">${actionButton}</div>
                      ${!isAvailable && eq.borrower ? `<div style="font-size:0.8em;color:#666;margin-top:3px;text-align:center;">👤 ${eq.borrower} | 📅 ${formatDateTime(eq.dt_borrow) || '未設定'}<br>⏰ ${formatDateTime(eq.dt_due) || '未設定'}</div>` : ''}
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  });

  list.innerHTML = html;
}

// 切換保管人分組展開/收起
function toggleKeeperGroup(headerEl) {
  const arrowEl = headerEl.querySelector('.keeper-arrow');
  const wrapperEl = headerEl.nextElementSibling;
  
  if (wrapperEl && arrowEl) {
    const isExpanded = wrapperEl.style.display !== 'none';
    
    if (isExpanded) {
      // 收起
      wrapperEl.style.display = 'none';
      arrowEl.style.transform = 'rotate(0deg)';
    } else {
      // 展開
      wrapperEl.style.display = 'block';
      arrowEl.style.transform = 'rotate(90deg)';
    }
  }
}

// =============================================
// 借用/歸還功能
// =============================================

// 開啟借用 Modal
function openBorrowModal(fixNo, deviceName, keeper) {
  const modal = document.getElementById('borrow-modal');
  const infoDiv = document.getElementById('borrow-equipment-info');
  
  if (infoDiv) {
    infoDiv.innerHTML = `
      <strong>設備編號：</strong>${fixNo}<br>
      <strong>設備名稱：</strong>${deviceName}<br>
      <strong>保管人：</strong>${keeper}
    `;
  }
  
  document.getElementById('borrow-fix-no').value = fixNo;
  
  // 設定最小日期為今天（台北時間）
  const now = new Date();
  const taipeiTime = new Date(now.getTime() + (8 * 60 * 60 * 1000));
  const taipeiDate = taipeiTime.toISOString().split('T')[0];
  document.getElementById('borrow-due-date').min = taipeiDate;
  
  if (modal) modal.style.display = 'block';
}

// 關閉借用 Modal
function closeBorrowModal() {
  const modal = document.getElementById('borrow-modal');
  if (modal) modal.style.display = 'none';
}

// 開啟歸還 Modal
function openReturnModal(fixNo, deviceName, borrower) {
  const modal = document.getElementById('return-modal');
  const infoDiv = document.getElementById('return-equipment-info');
  
  if (infoDiv) {
    infoDiv.innerHTML = `
      <strong>設備編號：</strong>${fixNo}<br>
      <strong>設備名稱：</strong>${deviceName}<br>
      <strong>借用人：</strong>${borrower}
    `;
  }
  
  document.getElementById('return-fix-no').value = fixNo;
  
  // 設定預設日期為今天（台北時間）
  const today = new Date();
  const taipeiTime = new Date(today.getTime() + (8 * 60 * 60 * 1000));
  const taipeiDate = taipeiTime.toISOString().split('T')[0];
  document.getElementById('return-date').value = taipeiDate;
  
  if (modal) modal.style.display = 'block';
}

// 關閉歸還 Modal
function closeReturnModal() {
  const modal = document.getElementById('return-modal');
  if (modal) modal.style.display = 'none';
}

// 借用設備
async function submitBorrow(formData) {
  try {
    // 使用 GET 請求避免 CORS preflight 問題
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'borrow');
    url.searchParams.append('fix_no', formData.fix_no);
    url.searchParams.append('borrower', formData.borrower);
    url.searchParams.append('dt_borrow', formData.dt_borrow);
    url.searchParams.append('dt_due', formData.dt_due);

    console.log('借用請求網址:', url.toString());

    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });

    const result = await res.json();

    if (result.success || result.status === 'success' || (!result.error && result.message)) {
      return { success: true, message: '✅ 借用成功！已通知保管人' };
    } else {
      throw new Error(result.error || '借用失敗');
    }
  } catch (err) {
    console.error('借用失敗:', err);
    return { success: false, message: `❌ ${err.message}` };
  }
}

// 歸還設備
async function submitReturn(formData) {
  try {
    // 使用 GET 請求避免 CORS preflight 問題
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'return');
    url.searchParams.append('fix_no', formData.fix_no);
    url.searchParams.append('dt_return', formData.dt_return);

    console.log('歸還請求網址:', url.toString());

    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });

    const result = await res.json();

    if (result.success || result.status === 'success' || (!result.error && result.message)) {
      return { success: true, message: '📧 歸還通知已發送！\n\n系統已寄信通知保管人（Keeper）\n請等待 Keeper 點擊郵件中的【確認已收到】按鈕後，設備狀態才會更新為「可借用」。' };
    } else {
      throw new Error(result.error || '歸還失敗');
    }
  } catch (err) {
    console.error('歸還失敗:', err);
    return { success: false, message: `❌ ${err.message}` };
  }
}

// 確認歸還（Keeper 使用）- 直接在主頁確認，不需要 email 驗證
async function confirmReturn(fixNo, deviceName, keeper) {
  if (!confirm(`確認歸還設備？\n\n設備編號：${fixNo}\n設備名稱：${deviceName}\n保管人：${keeper}\n\n確認後設備狀態將改為「可借用」。`)) {
    return;
  }
  
  const contentDiv = document.getElementById('content');
  if (contentDiv) {
    contentDiv.innerHTML = '<p class="loading">正在處理，請稍候...</p>';
  }
  
  try {
    // 使用 GET 請求，不需要 email 驗證
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'confirmReturn');
    url.searchParams.append('fix_no', fixNo);
    // 不再需要 keeper_email 參數

    console.log('確認歸還請求網址:', url.toString());

    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });

    const result = await res.json();

    if (result.success || result.status === 'success' || (!result.error && result.message)) {
      alert('✅ 歸還已確認！設備狀態已更新。');
      searchEquipment();  // 重新整理列表
    } else {
      throw new Error(result.error || '確認失敗');
    }
  } catch (err) {
    console.error('確認歸還失敗:', err);
    alert(`❌ 確認失敗：${err.message}`);
  }
}

// =============================================
// 登記功能
// =============================================

// 設備登記
async function registerEquipment(formData) {
  try {
    // 使用 GET 請求避免 CORS preflight 問題
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'register');
    url.searchParams.append('fix_type', formData.fix_type);
    url.searchParams.append('fix_no', formData.fix_no || '');
    url.searchParams.append('device_name', formData.device_name);
    url.searchParams.append('qty_asset', formData.qty_asset || '1');
    url.searchParams.append('keeper', formData.keeper || '');

    console.log('登記請求網址:', url.toString());

    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });

    const result = await res.json();

    if (result.success || result.status === 'success') {
      return { success: true, message: '✅ 設備登記成功！' };
    } else {
      throw new Error(result.error || '登記失敗');
    }
  } catch (err) {
    console.error('登記失敗:', err);
    return { success: false, message: `❌ ${err.message}` };
  }
}

// =============================================
// 初始化
// =============================================

document.addEventListener('DOMContentLoaded', () => {
  // 綁定搜尋按鈕
  const searchBtn = document.querySelector('.search-bar button');
  if (searchBtn) {
    searchBtn.addEventListener('click', searchEquipment);
  }

  // 綁定 Enter 鍵搜尋
  const searchInput = document.getElementById('search-keyword');
  if (searchInput) {
    searchInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        searchEquipment();
      }
    });
  }

  // 綁定借用表單
  const borrowForm = document.getElementById('borrow-form');
  if (borrowForm) {
    borrowForm.addEventListener('submit', async (e) => {
      e.preventDefault();

      const fixNo = document.getElementById('borrow-fix-no').value;
      const borrower = document.getElementById('borrow-name').value;
      const dtDue = document.getElementById('borrow-due-date').value;
      // 今天日期（台北時間）
      const now = new Date();
      const taipeiTime = new Date(now.getTime() + (8 * 60 * 60 * 1000));
      const dtBorrow = taipeiTime.toISOString().split('T')[0];

      const submitBtn = e.target.querySelector('button[type="submit"]');
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.textContent = '🔄 處理中...';
      }

      const result = await submitBorrow({ fix_no: fixNo, borrower: borrower, dt_borrow: dtBorrow, dt_due: dtDue });

      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = '確認借用';
      }

      alert(result.message);

      if (result.success) {
        closeBorrowModal();
        searchEquipment();
      }
    });
  }

  // 綁定歸還表單
  const returnForm = document.getElementById('return-form');
  if (returnForm) {
    returnForm.addEventListener('submit', async (e) => {
      e.preventDefault();

      const fixNo = document.getElementById('return-fix-no').value;
      const dtReturn = document.getElementById('return-date').value;

      const submitBtn = e.target.querySelector('button[type="submit"]');
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.textContent = '🔄 處理中...';
      }

      const result = await submitReturn({ fix_no: fixNo, dt_return: dtReturn });

      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = '確認歸還';
      }

      alert(result.message);

      if (result.success) {
        closeReturnModal();
        searchEquipment();
      }
    });
  }

  // 綁定登記表單
  const registerForm = document.getElementById('register-form');
  if (registerForm) {
    registerForm.addEventListener('submit', async (e) => {
      e.preventDefault();

      const formData = new FormData(e.target);
      const data = Object.fromEntries(formData);

      const submitBtn = e.target.querySelector('button[type="submit"]');
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.textContent = '🔄 登記中...';
      }

      const result = await registerEquipment(data);

      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = '登記設備';
      }

      alert(result.message);

      if (result.success) {
        e.target.reset();
        // 切換到設備列表分頁並重新查詢
        const equipmentTab = document.querySelector('[data-tab="equipment"]');
        if (equipmentTab) {
          equipmentTab.click();
          searchEquipment();
        }
      }
    });
  }

  // 頁面載入時自動查詢
  searchEquipment();
  
  // Modal 點擊外部關閉
  window.addEventListener('click', (e) => {
    const borrowModal = document.getElementById('borrow-modal');
    const returnModal = document.getElementById('return-modal');
    if (e.target === borrowModal) closeBorrowModal();
    if (e.target === returnModal) closeReturnModal();
  });
});

// =============================================
// 歷史紀錄功能
// =============================================

// 搜尋歷史紀錄
async function searchHistory() {
  const keyword = document.getElementById('history-keyword')?.value.trim() || '';
  const actionFilter = document.getElementById('history-action')?.value || '';
  const sortOrder = document.getElementById('history-sort')?.value || 'newest';  // 預設由新到舊

  const listEl = document.getElementById('history-list');
  if (listEl) {
    listEl.innerHTML = '<p style="text-align:center;color:#666;padding:40px;">🔄 載入中...</p>';
  }

  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'history');
    if (keyword) url.searchParams.append('keyword', keyword);
    if (actionFilter) url.searchParams.append('actionType', actionFilter);  // 修正：使用 actionType 避免衝突

    console.log('歷史紀錄請求網址:', url.toString());

    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });

    const data = await res.json();
    console.log('歷史紀錄回應:', data);

    if (data.error) {
      throw new Error(data.error);
    }

    const history = Array.isArray(data) ? data : (data.data || data.result || data.items || []);
    renderHistory(history, sortOrder);  // 傳遞排序參數
  } catch (err) {
    console.error('查詢歷史紀錄失敗:', err);
    if (listEl) {
      listEl.innerHTML = `<p style="text-align:center;color:#c00;padding:40px;">❌ 查詢失敗：${err.message}</p>`;
    }
  }
}

// 渲染歷史紀錄 - 按設備編號 + 借用人分組
function renderHistory(history, sortOrder = 'newest') {
  const list = document.getElementById('history-list');
  if (!list) return;

  if (!history || history.length === 0) {
    list.innerHTML = '<p style="text-align:center;color:#666;padding:40px;">目前沒有歷史紀錄</p>';
    return;
  }

  // 根據排序選項排序
  console.log('排序前:', history.length, '筆紀錄');
  console.log('第一筆時間戳:', history[0]?.timestamp);
  console.log('最後一筆時間戳:', history[history.length-1]?.timestamp);
  
  history.sort((a, b) => {
    const dateA = new Date(a.timestamp || 0);
    const dateB = new Date(b.timestamp || 0);
    const result = sortOrder === 'newest' ? (dateB - dateA) : (dateA - dateB);
    return result;
  });
  
  console.log('排序後（' + sortOrder + '）:', history[0]?.timestamp, '到', history[history.length-1]?.timestamp);

  // 按設備編號分組
  const deviceGroups = {};
  history.forEach(record => {
    const fixNo = record.fix_no || '無編號';
    if (!deviceGroups[fixNo]) {
      deviceGroups[fixNo] = {
        fix_no: fixNo,
        device_name: record.device_name,
        users: [],
        lastTimestamp: record.timestamp || ''  // 記錄該設備的最新時間戳
      };
    }
    
    // 更新最新時間戳（如果這筆紀錄更新）
    const recordTime = new Date(record.timestamp || 0).getTime();
    const currentLastTime = new Date(deviceGroups[fixNo].lastTimestamp || 0).getTime();
    if (recordTime > currentLastTime) {
      deviceGroups[fixNo].lastTimestamp = record.timestamp;
    }
    
    // 檢查是否為新的借用週期（borrow 動作）
    if (record.action === 'borrow') {
      const borrower = record.borrower || '未知';
      deviceGroups[fixNo].users.push({
        borrower: borrower,
        keeper: record.keeper,
        dt_borrow: record.dt_borrow,
        dt_due: record.dt_due,
        dt_return: '',
        return_confirmed: false,
        records: [record]
      });
    } else {
      // return 或 confirm，歸到最後一個借用週期
      const users = deviceGroups[fixNo].users;
      if (users.length > 0) {
        const lastUser = users[users.length - 1];
        lastUser.records.push(record);
        // 從 return/confirm 記錄中讀取正確的借用日期和預計歸還
        if (record.action === 'return' || record.action === 'confirm') {
          // 後端已修復：return/confirm 記錄的 dt_borrow=借用日期, dt_due=預計歸還, dt_confirmed=歸還日期
          if (record.dt_borrow) lastUser.dt_borrow = record.dt_borrow;
          if (record.dt_due) lastUser.dt_due = record.dt_due;
          // 後端回傳的是 dt_confirmed，不是 dt_return
          if (record.dt_return) lastUser.dt_return = record.dt_return;
          if (record.return_confirmed) lastUser.return_confirmed = record.return_confirmed;
        }
      }
    }
  });

  let html = '';
  
  // 將設備組按最新時間戳排序
  const sortedDeviceKeys = Object.keys(deviceGroups).sort((a, b) => {
    const timeA = new Date(deviceGroups[a].lastTimestamp || 0).getTime();
    const timeB = new Date(deviceGroups[b].lastTimestamp || 0).getTime();
    return sortOrder === 'newest' ? (timeB - timeA) : (timeA - timeB);
  });
  
  console.log('設備排序（' + sortOrder + '）:', sortedDeviceKeys.map(k => k + ':' + deviceGroups[k].lastTimestamp));
  
  sortedDeviceKeys.forEach((fixNo, deviceIndex) => {
    const group = deviceGroups[fixNo];
    const deviceExpanded = deviceIndex === 0; // 第一個設備預設展開
    
    html += `
      <div class="history-device-group" style="margin-bottom:20px;">
        <div class="history-device-header" onclick="toggleHistoryDevice(this)" style="cursor:pointer;user-select:none;background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);color:white;padding:12px 15px;border-radius:6px;margin-bottom:10px;font-weight:bold;font-size:1.1em;display:flex;align-items:center;">
          <span class="device-arrow" style="display:inline-block;width:12px;margin-right:8px;transition:transform 0.2s;${deviceExpanded ? 'transform:rotate(90deg)' : ''}">▶</span>
          <span>📦 ${fixNo} - ${group.device_name || '未知設備'}</span>
        </div>
        <div class="history-device-content" style="${deviceExpanded ? 'display:block;' : 'display:none;'}">
    `;
    
    // 每個借用人的紀錄
    group.users.forEach((user, userIndex) => {
      const hasConfirm = user.records.some(r => r.action === 'confirm' || r.action === 'confirmed');
      const hasReturn = user.records.some(r => r.action === 'return');
      const isExpanded = userIndex === 0; // 第一個預設展開
      
      // 判斷狀態
      let statusIcon, statusText;
      if (hasConfirm) {
        statusIcon = '✅';
        statusText = '已歸還';
      } else if (hasReturn) {
        statusIcon = '📥';
        statusText = '歸還（待確認）';
      } else {
        statusIcon = '📤';
        statusText = '借用';
      }
      
      html += `
        <div class="history-borrow-cycle" style="margin-bottom:10px;">
          <div class="history-borrow-header" onclick="toggleHistoryBorrow(event, this)" style="cursor:pointer;user-select:none;display:flex;align-items:center;padding:10px;background:#f8f9fa;border-radius:6px;border-left:4px solid #667eea;">
            <span class="borrow-arrow" style="display:inline-block;width:12px;margin-right:8px;transition:transform 0.2s;${isExpanded ? 'transform:rotate(90deg)' : ''}">▶</span>
            <span style="font-weight:bold;font-size:0.95em;">---> ${getAvatarHtml(user.borrower, 24)} ${user.borrower} ${statusIcon} ${statusText}</span>
          </div>
          <div class="history-borrow-detail" style="${isExpanded ? 'display:block;' : 'display:none;'}margin-left:20px;margin-top:8px;padding:10px;background:#fff;border-radius:6px;">
            <div style="font-size:0.9em;color:#666;line-height:1.8;">
              <div>借用日期：${formatDateTime(user.dt_borrow) || '－'}</div>
              <div>預計歸還：${formatDateTime(user.dt_due) || '－'}</div>
              ${user.dt_return ? `
                <div>${user.return_confirmed ? '歸還完成' : '歸還日期'}：${formatDateTime(user.dt_return)}${user.return_confirmed ? '' : '（待確認）'}</div>
              ` : '<div>歸還完成：－</div>'}
            </div>
          </div>
        </div>
      `;
    });
    
    html += `</div></div>`;
  });
  
  list.innerHTML = html;
}

// 切換歷史紀錄設備展開/收起
function toggleHistoryDevice(headerEl) {
  const arrowEl = headerEl.querySelector('.device-arrow');
  const contentEl = headerEl.nextElementSibling;
  
  if (contentEl && arrowEl) {
    const isExpanded = contentEl.style.display !== 'none';
    
    if (isExpanded) {
      // 收起
      contentEl.style.display = 'none';
      arrowEl.style.transform = 'rotate(0deg)';
    } else {
      // 展開
      contentEl.style.display = 'block';
      arrowEl.style.transform = 'rotate(90deg)';
    }
  }
}

// 切換歷史紀錄借用週期展開/收起
function toggleHistoryBorrow(event, headerEl) {
  event.stopPropagation(); // 防止事件冒泡到設備標題
  
  const arrowEl = headerEl.querySelector('.borrow-arrow');
  const detailEl = headerEl.nextElementSibling;
  
  if (detailEl && arrowEl) {
    const isExpanded = detailEl.style.display !== 'none';
    
    if (isExpanded) {
      // 收起
      detailEl.style.display = 'none';
      arrowEl.style.transform = 'rotate(0deg)';
    } else {
      // 展開
      detailEl.style.display = 'block';
      arrowEl.style.transform = 'rotate(90deg)';
    }
  }
}

// 切換歷史紀錄週期展開/收起（舊函數，保留相容性）
function toggleHistoryCycle(headerEl) {
  const arrowEl = headerEl.querySelector('.cycle-arrow');
  const detailEl = headerEl.nextElementSibling;
  
  if (detailEl && arrowEl) {
    const isExpanded = detailEl.style.display !== 'none';
    
    if (isExpanded) {
      // 收起
      detailEl.style.display = 'none';
      arrowEl.style.transform = 'rotate(0deg)';
    } else {
      // 展開
      detailEl.style.display = 'block';
      arrowEl.style.transform = 'rotate(90deg)';
    }
  }
}

// 綁定歷史紀錄搜尋
const historySearchBtn = document.querySelector('#history-tab .search-bar button');
if (historySearchBtn) {
  historySearchBtn.addEventListener('click', searchHistory);
}

const historyInput = document.getElementById('history-keyword');
if (historyInput) {
  historyInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      searchHistory();
    }
  });
}

// 注意：排序選單事件綁定移到分頁切換時（因為元素預設隱藏）
// 綁定歷史排序選單變更事件 - 在分頁切換到 history 時綁定

// 綁定分頁切換
document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', function() {
    const tab = this.dataset.tab;
    
    // 移除所有 active
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    // 添加 active
    this.classList.add('active');
    document.getElementById(`${tab}-tab`).classList.add('active');
    
    // 如果切換到歷史紀錄分頁，自動載入並綁定排序事件
    if (tab === 'history') {
      searchHistory();
      
      // 綁定排序選單事件（只在切換到歷史頁時綁定一次）
      const historySortSelect = document.getElementById('history-sort');
      if (historySortSelect && !historySortSelect._hasEventListener) {
        historySortSelect.addEventListener('change', () => {
          searchHistory();  // 排序變更時自動重新搜尋
        });
        historySortSelect._hasEventListener = true;  // 標記已綁定，避免重複
      }
    }
    // 如果切換到個人設定分頁，載入頭像列表
    if (tab === 'settings') {
      loadAvatarList();
    }
  });
});

// =============================================
// 頭像功能
// =============================================

// 頭像本地快取（格式：{ "姓名": "URL" }）
let avatarCache = {};

/**
 * 載入頭像快取（從 localStorage）
 */
async function loadAvatarCache() {
  try {
    // 嘗試從 localStorage 載入
    const cached = localStorage.getItem('avatarCache');
    if (cached) {
      avatarCache = JSON.parse(cached);
      console.log('從 localStorage 載入頭像快取:', Object.keys(avatarCache).length, '個');
    }
  } catch (err) {
    console.log('載入頭像快取失敗:', err);
  }
}

/**
 * 儲存頭像快取到 localStorage
 */
function saveAvatarCache() {
  try {
    localStorage.setItem('avatarCache', JSON.stringify(avatarCache));
  } catch (err) {
    console.log('儲存頭像快取失敗:', err);
  }
}

/**
 * 取得頭像 HTML（圖片或預設 emoji）
 */
function getAvatarHtml(name, size = 55) {  // 圖片預設 55px
  if (!name) return `<span style="font-size:30px;">👤</span>`;  // Emoji 固定 30px
  
  const url = avatarCache[name];
  if (url) {
    return `<img src="${url}" style="width:${size}px;height:${size}px;border-radius:50%;object-fit:cover;vertical-align:middle;" onerror="this.style.display='none';this.nextElementSibling.style.display='inline';"><span style="font-size:30px;display:none;">👤</span>`;
  }
  return `<span style="font-size:30px;">👤</span>`;
}

/**
 * 載入所有頭像列表（從 Google Sheet）
 */
async function loadAvatarList() {
  const listEl = document.getElementById('avatar-list');
  if (!listEl) return;
  
  listEl.innerHTML = '<p style="text-align:center;color:#666;padding:20px;">🔄 載入中...</p>';
  
  try {
    // 從 GAS 取得頭像列表
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'getAvatarList');
    
    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });
    
    const data = await res.json();
    const avatars = Array.isArray(data) ? data : (data.data || data.result || []);
    
    if (!avatars || avatars.length === 0) {
      listEl.innerHTML = '<p style="color:#888;">目前沒有已上傳的頭像</p>';
      return;
    }
    
    // 更新本地快取
    avatars.forEach(item => {
      if (item.name && item.avatar_url) {
        avatarCache[item.name] = item.avatar_url;
      }
    });
    saveAvatarCache();
    
    // 顯示頭像列表
    let html = '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(100px,1fr));gap:15px;">';
    for (const [name, url] of Object.entries(avatarCache)) {
      html += `
        <div style="text-align:center;">
          <img src="${url}" style="width:60px;height:60px;border-radius:50%;object-fit:cover;border:2px solid #ddd;">
          <div style="font-size:0.85em;margin-top:5px;">${name}</div>
        </div>
      `;
    }
    html += '</div>';
    listEl.innerHTML = html;
  } catch (err) {
    console.error('載入頭像列表失敗:', err);
    listEl.innerHTML = `<p style="color:red;">❌ 載入失敗：${err.message}</p>`;
  }
}

/**
 * 壓縮圖片並轉為 base64
 */
function compressImage(file, maxWidth = 100, quality = 0.6) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      const img = new Image();
      img.onload = function() {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;
        
        // 等比例縮小
        if (width > maxWidth) {
          height = Math.round(height * maxWidth / width);
          width = maxWidth;
        }
        
        canvas.width = width;
        canvas.height = height;
        
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);
        
        // 輸出為 JPEG（壓縮率更高）
        const compressed = canvas.toDataURL('image/jpeg', quality);
        resolve(compressed);
      };
      img.onerror = function(e) {
        console.error('圖片載入失敗:', e);
        reject(new Error('無法解碼圖片，請嘗試其他圖片格式（JPG/PNG）'));
      };
      img.crossOrigin = 'anonymous';
      img.src = e.target.result;
    };
    reader.onerror = function() {
      reject(new Error('讀取檔案失敗'));
    };
    reader.readAsDataURL(file);
  });
}

/**
 * 上傳頭像到 Google Drive（使用 GET，壓縮到 150x150）
 */
async function uploadAvatar(name, file) {
  try {
    // 壓縮圖片到 150x150，品質 0.8（更清楚）
    const compressedData = await compressImage(file, 150, 0.8);
    
    console.log('頭像上傳開始，data URL 長度:', compressedData.length);
    
    // 使用 GET 請求傳送
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'uploadAvatar');
    url.searchParams.append('user_name', name);
    url.searchParams.append('image_data', compressedData);
    
    console.log('GET URL 長度:', url.toString().length);
    
    // 檢查 URL 是否太長（超過 8000 字元可能會失敗）
    if (url.toString().length > 8000) {
      console.warn('URL 太長，可能失敗，嘗試再次壓縮...');
      const recompressed = await compressImage(file, 80, 0.5);
      url.searchParams.set('image_data', recompressed);
      console.log('再次壓縮後 URL 長度:', url.toString().length);
    }
    
    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });
    
    if (!res.ok) {
      const errorText = await res.text();
      throw new Error('HTTP ' + res.status + ': ' + errorText);
    }
    
    const result = await res.json();
    
    if (result.success || result.url) {
      // 更新本地快取
      avatarCache[name] = result.url;
      saveAvatarCache();
      return result;
    } else {
      throw new Error(result.error || '上傳失敗');
    }
  } catch (err) {
    console.error('上傳頭像失敗:', err);
    throw err;  // 重新拋出讓調用者知道錯誤
  }
}

// 初始化頭像功能
loadAvatarCache();

// 綁定頭像表單
const avatarForm = document.getElementById('avatar-form');
if (avatarForm) {
  // 圖片預覽
  const avatarFile = document.getElementById('avatar-file');
  const avatarPreview = document.getElementById('avatar-preview');
  
  if (avatarFile && avatarPreview) {
    avatarFile.addEventListener('change', async function(e) {
      const file = e.target.files[0];
      if (file) {
        // 限制圖片大小，最大 500KB
        if (file.size > 2 * 1024 * 1024) {
          alert('圖片太大，請選擇小於 2MB 的圖片');
          e.target.value = '';
          return;
        }
        
        try {
          // 壓縮並顯示預覽
          const compressed = await compressImage(file, 150, 0.7);
          avatarPreview.src = compressed;
          avatarPreview.style.display = 'block';
        } catch (err) {
          alert('圖片處理失敗: ' + err.message);
        }
      }
    });
  }
  
  // 表單提交
  avatarForm.addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const name = document.getElementById('avatar-name').value.trim();
    const file = document.getElementById('avatar-file').files[0];
    const statusEl = document.getElementById('avatar-status');
    const submitBtn = avatarForm.querySelector('button[type="submit"]');
    
    if (!name) {
      alert('請輸入姓名');
      return;
    }
    
    if (!file) {
      alert('請選擇圖片');
      return;
    }
    
    submitBtn.disabled = true;
    submitBtn.textContent = '🔄 上傳中...';
    statusEl.innerHTML = '';
    
    try {
      const result = await uploadAvatar(name, file);
      statusEl.innerHTML = '<p style="color:green;">✅ 頭像上傳成功！</p>';
      avatarPreview.src = result.url;
      avatarPreview.style.display = 'block';
      loadAvatarList(); // 更新頭像列表
    } catch (err) {
      statusEl.innerHTML = `<p style="color:red;">❌ ${err.message}</p>`;
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = '上傳頭像';
    }
  });
}
