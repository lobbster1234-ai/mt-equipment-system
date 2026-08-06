// =============================================
// MT 設備系統 - Google Apps Script 前端
// =============================================

const GAS_URL = 'https://script.google.com/macros/s/AKfycbxeI5xC33a6Ry634g6kwBPK9feElH_tTPtQYeWcH4ReiEiiq5I9yIetv8ugAFDgJkHh1A/exec';

// 載入中動畫（旋轉圈圈 + 文字）
function loadingHtml() {
  return '<div class="loading-box"><div class="spinner"></div><span>載入中...</span></div>';
}

// 統一的 GAS GET 請求，處理「GAS 轉址不穩、回傳 HTML 而非 JSON」的情況。
// - 解析成功 → 回傳 JSON 物件
// - 讀取類（傳 { retries: N }）→ 收到 HTML 時自動重試（讀取為冪等，重試安全）
// - 仍失敗 → 丟出帶 isGasGlitch=true 的錯誤，讓呼叫端顯示友善提示而非 Unexpected token 錯誤
async function gasGetJson(url, opts) {
  opts = opts || {};
  const retries = opts.retries || 0;
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(url, { method: 'GET', redirect: 'follow' });
    const text = await res.text();
    try {
      return JSON.parse(text);
    } catch (e) {
      if (attempt < retries) { await new Promise(r => setTimeout(r, 1200)); continue; }
      console.warn('GAS 回應非 JSON（轉址不穩）:', text.slice(0, 120));
      const err = new Error('伺服器回應不穩定，請稍後再試');
      err.isGasGlitch = true;
      throw err;
    }
  }
}

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

  // 是否為「無條件」的預設查詢（頁面初次載入即為此）
  const isDefaultQuery = !keyword && !department && !status;

  const listEl = document.getElementById('equipment-list');
  let shownCache = false;
  if (listEl) {
    // 預設查詢且有快取時，先秒顯示上次的資料，背景再抓最新的來更新
    if (isDefaultQuery) {
      try {
        const cached = localStorage.getItem('mt_equipment_cache');
        if (cached) {
          const arr = JSON.parse(cached);
          if (Array.isArray(arr) && arr.length) {
            renderEquipment(arr);
            shownCache = true;
          }
        }
      } catch (e) { /* 快取解析失敗就忽略，改用轉圈圈 */ }
    }
    if (!shownCache) listEl.innerHTML = loadingHtml();
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
    // 快取預設查詢結果，供下次開啟時秒顯示
    if (isDefaultQuery) {
      try { localStorage.setItem('mt_equipment_cache', JSON.stringify(equipment)); } catch (e) { /* 空間不足等狀況忽略 */ }
    }
  } catch (err) {
    console.error('查詢失敗:', err);
    // 若已先顯示快取資料，背景更新失敗就保留舊資料，不覆蓋成錯誤畫面
    if (listEl && !shownCache) {
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
                // 標準化狀態值（處理大小寫和空格）
                const rawStatus = (eq.status || '').toString().trim().toLowerCase();
                const isAvailable = rawStatus === 'available' || rawStatus === '可借用' || rawStatus === '';
                const isBorrowed = rawStatus === 'borrowed' || rawStatus === '已借出' || rawStatus === '借用中';
                const isBorrowPending = rawStatus === 'borrow_pending' || rawStatus === '借用審核中';
                const isReturnPending = rawStatus === 'return_pending' || rawStatus === '歸還認證中';
                const isConfirmed = eq.return_confirmed === true || eq.return_confirmed === 'true' || eq.return_confirmed === 1;
                
                // 除錯：記錄每個設備的狀態
                console.log(`設備 ${eq.fix_no} 狀態: "${eq.status}" (raw: "${rawStatus}") => available=${isAvailable}, borrowed=${isBorrowed}, borrowPending=${isBorrowPending}`);
                
                let statusHtml;
                if (isAvailable) {
                  statusHtml = '<span style="color:#0a0;">✅ 可借用</span>';
                } else if (isBorrowed) {
                  statusHtml = '<span style="color:#c00;">📤 借用中</span>';
                } else if (isBorrowPending) {
                  statusHtml = '<span style="color:#ffc107;">⏳ 借用審核中</span>';
                } else if (isReturnPending) {
                  statusHtml = '<span style="color:#17a2b8;">⏳ 歸還認證中</span>';
                } else if (isConfirmed) {
                  statusHtml = '<span style="color:#999;">✅ 已確認</span>';
                } else {
                  // 其他情況（可能是借用中但未明確標記）
                  statusHtml = '<span style="color:#c00;">📤 借用中</span>';
                }
                
                // 借用/歸還/延後按鈕
                let actionButton = '';
                // URL 編碼設備名稱，避免特殊字元破壞 HTML/JavaScript 語法
                const encodedDeviceName = encodeURIComponent(eq.device_name || '');
                const encodedBorrower = encodeURIComponent(eq.borrower || '');
                const encodedKeeper = encodeURIComponent(eq.keeper || '');
                if (isAvailable) {
                  actionButton = `<button class="btn-borrow-sm" onclick="openBorrowModal('${eq.fix_no}', decodeURIComponent('${encodedDeviceName}'), decodeURIComponent('${encodedKeeper}'))">借用</button>`;
                } else if (isBorrowed) {
                  // 借用中 - 顯示【歸還】和【延後】按鈕
                  actionButton = `
                    <button class="btn-postpone-sm" onclick="openPostponeModal('${eq.fix_no}', decodeURIComponent('${encodedDeviceName}'), decodeURIComponent('${encodedBorrower}'), '${eq.dt_due}')">⏰ 續借</button>
                    <button class="btn-return-sm" onclick="openReturnModal('${eq.fix_no}', decodeURIComponent('${encodedDeviceName}'), decodeURIComponent('${encodedBorrower}'))" style="margin-left:5px;">📧 歸還</button>
                  `;
                } else if (isBorrowPending) {
                  // 借用審核中，顯示提示文字
                  actionButton = '<span style="color:#ffc107;font-size:0.85em;">⏳ 等待 Keeper 審核</span>';
                } else if (isReturnPending) {
                  // 歸還認證中，不顯示按鈕
                  actionButton = '<span style="color:#17a2b8;font-size:0.85em;">等待 Keeper 確認</span>';
                } else if (!isConfirmed && !isAvailable) {
                  // 借用中，未歸還（其他狀態）
                  const hasReturnDate = eq.dt_return && eq.dt_return.toString().trim() !== '';
                  if (hasReturnDate) {
                    actionButton = '<span style="color:#17a2b8;font-size:0.85em;">⏳ 待確認</span>';
                  } else {
                    actionButton = `<button class="btn-return-sm" onclick="openReturnModal('${eq.fix_no}', decodeURIComponent('${encodedDeviceName}'), decodeURIComponent('${encodedBorrower}'))">📧 歸還</button>`;
                  }
                } else {
                  actionButton = '<span style="color:#999;font-size:0.85em;">已確認</span>';
                }
                
                return `
                  <tr>
                    <td>${escapeHtml(eq.fix_type || '')}</td>
                    <td>${escapeHtml(eq.fix_no || '')}</td>
                    <td>${escapeHtml(eq.device_name || '')}</td>
                    <td>${escapeHtml(eq.qty_asset || '1')}</td>
                    <td>
                      ${statusHtml}
                      <div style="margin-top:5px;">${actionButton}</div>
                      ${(isBorrowed || isBorrowPending || isReturnPending || (!isAvailable && !isBorrowed && !isBorrowPending && !isReturnPending && !isConfirmed)) && eq.borrower ? `<div style="font-size:0.8em;color:#666;margin-top:3px;text-align:center;">👤 ${escapeHtml(eq.borrower)} | 📅 ${formatDateTime(eq.dt_borrow) || '未設定'}<br>⏰ ${formatDateTime(eq.dt_due) || '未設定'}</div>` : ''}
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

/**
 * 從 GAS 根據姓名查詢 Email
 */
async function fetchEmailByName(name) {
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'getEmailByName');
    url.searchParams.append('name', name);
    
    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });
    
    const result = await res.json();
    if (result.success && result.email) {
      return result.email;
    }
    return '';
  } catch (err) {
    console.error('查詢 email 失敗:', err);
    return '';
  }
}

/**
 * 處理借用表單提交（繞過 form submit 事件）
 */
async function handleBorrowSubmit() {
  console.log('=== handleBorrowSubmit 被呼叫 ===');
  
  const fixNo = document.getElementById('borrow-fix-no').value;
  const borrower = document.getElementById('borrow-name').value;
  const borrowerEmail = document.getElementById('borrow-email')?.value?.trim() || '';
  const dtDue = document.getElementById('borrow-due-date').value;
  
  console.log('fixNo:', fixNo);
  console.log('借用人姓名:', borrower);
  console.log('借用人 email:', borrowerEmail);
  console.log('預計歸還:', dtDue);
  
  // 檢查必填欄位
  if (!borrower) {
    alert('請填寫借用人姓名');
    return;
  }
  if (!borrowerEmail) {
    alert('請填寫電子郵件（用於接收審核結果通知）');
    return;
  }
  if (!dtDue) {
    alert('請選擇預計歸還日期時間');
    return;
  }
  
  // 借用日期：四捨五入到最接近的整點
  const now = new Date();
  const hours = now.getHours();
  const minutes = now.getMinutes();
  let borrowHour = hours;
  if (minutes >= 30) {
    borrowHour = hours + 1; // >=30分鐘，小時+1
  }
  // 處理跨日
  let borrowDay = now.getDate();
  if (borrowHour >= 24) {
    borrowHour = 0;
    borrowDay++;
  }
  // 格式化為 yyyy-MM-ddTHH:mm
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(borrowDay).padStart(2, '0');
  const hour = String(borrowHour).padStart(2, '0');
  const dtBorrow = `${year}-${month}-${day}T${hour}:00`;
  
  // 判斷是否為訪客
  const user = JSON.parse(localStorage.getItem('mt_user') || '{}');
  const isGuest = user.role !== 'admin';
  
  const submitBtn = document.querySelector('#borrow-form button');
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.textContent = '🔄 處理中...';
  }
  
  let result;
  try {
    console.log('開始呼叫借用 API...');
    if (isGuest) {
      console.log('訪客模式，呼叫 requestBorrow');
      result = await requestBorrow({ 
        fix_no: fixNo, 
        borrower: borrower, 
        borrower_email: borrowerEmail,
        dt_borrow: dtBorrow, 
        dt_due: dtDue 
      });
    } else {
      console.log('管理員模式，呼叫 submitBorrow');
      result = await submitBorrow({ 
        fix_no: fixNo, 
        borrower: borrower, 
        borrower_email: borrowerEmail,
        dt_borrow: dtBorrow, 
        dt_due: dtDue 
      });
    }
    console.log('API 回應:', result);
  } catch (err) {
    console.error('借用 API 錯誤:', err);
    alert('借用失敗：' + err.message);
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.textContent = '確認借用';
    }
    return;
  }
  
  if (submitBtn) {
    submitBtn.disabled = false;
    submitBtn.textContent = '確認借用';
  }
  
  alert(result.message);
  
  if (result.success) {
    closeBorrowModal();
    searchEquipment();
  }
}

// 開啟借用 Modal
function openBorrowModal(fixNo, deviceName, keeper) {
  console.log('=== openBorrowModal 被呼叫 ===');
  console.log('fixNo:', fixNo, 'deviceName:', deviceName, 'keeper:', keeper);
  
  const modal = document.getElementById('borrow-modal');
  const infoDiv = document.getElementById('borrow-equipment-info');
  
  if (!modal) {
    console.error('找不到 borrow-modal 元素');
    return;
  }
  if (!infoDiv) {
    console.error('找不到 borrow-equipment-info 元素');
    return;
  }
  
  if (infoDiv) {
    infoDiv.innerHTML = `
      <strong>設備編號：</strong>${escapeHtml(fixNo)}<br>
      <strong>設備名稱：</strong>${escapeHtml(deviceName)}<br>
      <strong>保管人：</strong>${escapeHtml(keeper)}
    `;
  }
  
  const fixNoInput = document.getElementById('borrow-fix-no');
  if (!fixNoInput) {
    console.error('找不到 borrow-fix-no 輸入框');
    return;
  }
  fixNoInput.value = fixNo;
  
  // 自動填入登入者資訊
  const user = JSON.parse(localStorage.getItem('mt_user'));
  console.log('目前使用者:', user);
  
  const borrowNameInput = document.getElementById('borrow-name');
  const borrowEmailGroup = document.getElementById('borrow-email-group');
  const borrowEmailInput = document.getElementById('borrow-email');
  
  // 確保 email 欄位顯示
  if (borrowEmailGroup) {
    borrowEmailGroup.style.display = 'block';
  }
  
  // 清除之前的值
  if (borrowNameInput) {
    borrowNameInput.value = '';
    borrowNameInput.readOnly = false;
    borrowNameInput.style.background = '#fff';
  }
  if (borrowEmailInput) {
    borrowEmailInput.value = '';
    borrowEmailInput.readOnly = false;
  }
  
  if (user && user.role === 'admin' && borrowNameInput) {
    // 管理員登入，自動填入姓名和 email（姓名寫死，email 可編輯）
    console.log('管理員模式：自動填入姓名和 email');
    borrowNameInput.value = user.name || '';
    borrowNameInput.readOnly = true;
    borrowNameInput.style.background = '#e9ecef';
    borrowNameInput.style.cursor = 'not-allowed';
    
    // 先從 localStorage 即時取得 email（可編輯）
    if (borrowEmailInput) {
      borrowEmailInput.value = user.email || '';
      borrowEmailInput.readOnly = false;
      borrowEmailInput.style.background = '#fff';
      borrowEmailInput.style.cursor = 'text';
      console.log('已從 localStorage 填入 email:', user.email);
    }
    
    // 背景更新：在背景從 GAS 更新 email（非同步，不阻擋 UI）
    if (user.name) {
      fetchEmailByName(user.name).then(email => {
        if (email) {
          // 如果 GAS 回傳的 email，更新 localStorage 和輸入框
          user.email = email;
          localStorage.setItem('mt_user', JSON.stringify(user));
          // 如果輸入框還是空的或是預設值，才更新
          if (!borrowEmailInput.value || borrowEmailInput.value === user.email) {
            borrowEmailInput.value = email;
          }
          console.log('已更新 email:', email);
        }
      });
    }
  } else if (borrowNameInput) {
    // 訪客登入，不預填
    console.log('訪客模式');
    borrowNameInput.value = '';
  }
  
  // 設定最小日期時間為現在（台北時間），強制整點
  const now = new Date();
  const taipeiTime = new Date(now.getTime() + (8 * 60 * 60 * 1000));
  taipeiTime.setMinutes(0, 0, 0); // 強制整點
  const taipeiDateTime = taipeiTime.toISOString().slice(0, 16); // yyyy-MM-ddTHH:mm
  const borrowDueDateInput = document.getElementById('borrow-due-date');
  if (borrowDueDateInput) {
    borrowDueDateInput.min = taipeiDateTime;
    
    // 監聽變更事件，強制改為最接近的整點（四捨五入）- borrowDueDateInput
    borrowDueDateInput.addEventListener('change', function() {
      console.log('借用預計歸還時間變更:', this.value);
      if (this.value) {
        // datetime-local 格式: 2026-05-20T14:46 或 2026-05-20T23:36
        const parts = this.value.split('T');
        const datePart = parts[0];
        const timePart = parts[1];
        const hourPart = timePart.substring(0, 2);
        const minPart = timePart.substring(3, 5);
        
        const hours = parseInt(hourPart, 10);
        const minutes = parseInt(minPart, 10);
        
        console.log('解析 - 日期:', datePart, '時間:', timePart, '時:', hours, '分:', minutes);
        
        // 計算新的整點時間
        let newHours = hours;
        if (minutes >= 30) {
          newHours = hours + 1;
          console.log('進位到下一小時:', newHours);
        }
        
        // 處理跨日（例如 23:30 -> 00:00 隔天）
        let newDatePart = datePart;
        if (newHours >= 24) {
          newHours = 0;
          // 手動計算隔天日期（避免 Date 物件時區問題）
          const [y, m, d] = datePart.split('-').map(Number);
          let newDay = d + 1;
          let newMonth = m;
          let newYear = y;
          
          // 檢查是否跨月（例如 1/31 -> 2/1）
          const daysInMonth = new Date(y, m, 0).getDate();
          if (newDay > daysInMonth) {
            newDay = 1;
            newMonth = m + 1;
            // 檢查是否跨年
            if (newMonth > 12) {
              newMonth = 1;
              newYear = y + 1;
            }
          }
          
          const yStr = String(newYear);
          const mStr = String(newMonth).padStart(2, '0');
          const dStr = String(newDay).padStart(2, '0');
          newDatePart = `${yStr}-${mStr}-${dStr}`;
          console.log('跨日：', datePart, '->', newDatePart);
        }
        
        // 格式化新值
        const newHourStr = String(newHours).padStart(2, '0');
        const newValue = newDatePart + 'T' + newHourStr + ':00';
        console.log('修正後時間:', newValue);
        this.value = newValue;
      }
    });
  }
  
  if (modal) modal.style.display = 'flex';
}

// 關閉借用 Modal
function closeBorrowModal() {
  const modal = document.getElementById('borrow-modal');
  // 重置 email 欄位
  const borrowEmailGroup = document.getElementById('borrow-email-group');
  const borrowEmailInput = document.getElementById('borrow-email');
  if (borrowEmailGroup) borrowEmailGroup.style.display = 'none';
  if (borrowEmailInput) borrowEmailInput.value = '';
  if (modal) modal.style.display = 'none';
}

// 開啟歸還 Modal
function openReturnModal(fixNo, deviceName, borrower) {
  const modal = document.getElementById('return-modal');
  const infoDiv = document.getElementById('return-equipment-info');
  
  if (infoDiv) {
    infoDiv.innerHTML = `
      <strong>設備編號：</strong>${escapeHtml(fixNo)}<br>
      <strong>設備名稱：</strong>${escapeHtml(deviceName)}<br>
      <strong>借用人：</strong>${escapeHtml(borrower)}
    `;
  }
  
  document.getElementById('return-fix-no').value = fixNo;
  
  // 設定預設日期時間為現在（台北時間），四捨五入到最接近的整點
  const now = new Date();
  
  // 轉換為台北時間字串 (UTC+8)
  const taipeiOffset = 8 * 60 * 60 * 1000; // 8小時毫秒
  const taipeiTime = new Date(now.getTime() + taipeiOffset);
  
  // 取得台北時間的各個部分（注意：getUTCXXX 在加了 offset 後就是台北時間）
  const year = taipeiTime.getUTCFullYear();
  const month = taipeiTime.getUTCMonth() + 1; // 0-indexed
  const day = taipeiTime.getUTCDate();
  const hours = taipeiTime.getUTCHours();
  const minutes = taipeiTime.getUTCMinutes();
  
  console.log('現在台北時間:', year, month, day, hours, ':', minutes);
  
  // 四捨五入到最接近的整點
  let newHours = hours;
  if (minutes >= 30) {
    newHours = hours + 1;
  }
  
  // 處理跨日
  let newYear = year;
  let newMonth = month;
  let newDay = day;
  
  if (newHours >= 24) {
    newHours = 0;
    newDay++;
    
    // 檢查跨月
    const daysInMonth = new Date(year, month, 0).getDate();
    if (newDay > daysInMonth) {
      newDay = 1;
      newMonth++;
      // 檢查跨年
      if (newMonth > 12) {
        newMonth = 1;
        newYear++;
      }
    }
  }
  
  // 格式化
  const yStr = String(newYear);
  const mStr = String(newMonth).padStart(2, '0');
  const dStr = String(newDay).padStart(2, '0');
  const hStr = String(newHours).padStart(2, '0');
  const taipeiDateTime = `${yStr}-${mStr}-${dStr}T${hStr}:00`;
  
  console.log('歸還時間設定為:', taipeiDateTime);
  
  const returnDateInput = document.getElementById('return-date');
  if (returnDateInput) {
    // 如果是 type="date"，改成 type="datetime-local"
    if (returnDateInput.type === 'date') {
      returnDateInput.type = 'datetime-local';
    }
    returnDateInput.value = taipeiDateTime;
  }
  
  if (modal) modal.style.display = 'flex';
}

// 關閉歸還 Modal
function closeReturnModal() {
  const modal = document.getElementById('return-modal');
  if (modal) modal.style.display = 'none';
}

// 開啟延後歸還 Modal
function openPostponeModal(fixNo, deviceName, borrower, currentDueDate) {
  // 檢查是否已經有延後 Modal，如果沒有則動態建立
  let modal = document.getElementById('postpone-modal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'postpone-modal';
    modal.className = 'modal';
    modal.style.cssText = 'display:none;position:fixed;z-index:1000;left:0;top:0;width:100%;height:100%;overflow-y:auto;padding:20px;background-color:rgba(0,0,0,0.5);';
    modal.innerHTML = `
      <div style="background-color:#fefefe;margin:auto;padding:20px;border:1px solid #888;width:90%;max-width:500px;border-radius:12px;position:relative;">
        <span onclick="closePostponeModal()" style="color:#aaa;float:right;font-size:28px;font-weight:bold;cursor:pointer;">&times;</span>
        <h2 style="color:#667eea;margin-bottom:20px;">⏰ 續借</h2>
        <div id="postpone-info" style="background:#f8f9fa;padding:15px;border-radius:8px;margin-bottom:20px;"></div>
        <form id="postpone-form">
          <input type="hidden" id="postpone-fix-no">
          <div style="margin-bottom:15px;">
            <label style="display:block;margin-bottom:5px;font-weight:bold;">新的預計歸還時間：</label>
            <input type="datetime-local" id="postpone-new-due-date" step="3600" required style="width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;font-size:1em;">
          </div>
          <div style="text-align:center;margin-top:20px;">
            <button type="submit" class="btn" style="background:#ffc107;color:#000;padding:12px 24px;border:none;border-radius:8px;cursor:pointer;font-size:1em;">✅ 確認續借</button>
            <button type="button" onclick="closePostponeModal()" style="background:#6c757d;color:#fff;padding:12px 24px;border:none;border-radius:8px;cursor:pointer;font-size:1em;margin-left:10px;">取消</button>
          </div>
        </form>
      </div>
    `;
    document.body.appendChild(modal);
    
    // 綁定表單提交事件
    const form = document.getElementById('postpone-form');
    if (form) {
      form.addEventListener('submit', handlePostponeSubmit);
    }
    
    // 綁定 datetime-local 變更事件（四捨五入到整點）
    const dateInput = document.getElementById('postpone-new-due-date');
    if (dateInput) {
      dateInput.addEventListener('change', function() {
        if (this.value) {
          const parts = this.value.split('T');
          const datePart = parts[0];
          const timePart = parts[1];
          const hourPart = timePart.substring(0, 2);
          const minPart = timePart.substring(3, 5);
          
          const hours = parseInt(hourPart, 10);
          const minutes = parseInt(minPart, 10);
          
          let newHours = hours;
          if (minutes >= 30) {
            newHours = hours + 1;
          }
          
          // 處理跨日
          let newDatePart = datePart;
          if (newHours >= 24) {
            newHours = 0;
            const [y, m, d] = datePart.split('-').map(Number);
            let newDay = d + 1;
            let newMonth = m;
            let newYear = y;
            
            const daysInMonth = new Date(y, m, 0).getDate();
            if (newDay > daysInMonth) {
              newDay = 1;
              newMonth++;
              if (newMonth > 12) {
                newMonth = 1;
                newYear++;
              }
            }
            
            const yStr = String(newYear);
            const mStr = String(newMonth).padStart(2, '0');
            const dStr = String(newDay).padStart(2, '0');
            newDatePart = `${yStr}-${mStr}-${dStr}`;
          }
          
          const newHourStr = String(newHours).padStart(2, '0');
          const newValue = newDatePart + 'T' + newHourStr + ':00';
          this.value = newValue;
        }
      });
    }
  }
  
  // 設定當前資訊
  const infoDiv = document.getElementById('postpone-info');
  if (infoDiv) {
    infoDiv.innerHTML = `
      <strong>設備編號：</strong>${escapeHtml(fixNo)}<br>
      <strong>設備名稱：</strong>${escapeHtml(deviceName)}<br>
      <strong>借用人：</strong>${escapeHtml(borrower)}<br>
      <strong>目前預計歸還：</strong>${escapeHtml(currentDueDate || '未設定')}
    `;
  }
  
  document.getElementById('postpone-fix-no').value = fixNo;
  
  // 設定新的預計歸還時間預設值（從當前時間往後 1 天，四捨五入到整點）
  const now = new Date();
  const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  const taipeiOffset = 8 * 60 * 60 * 1000;
  const taipeiTime = new Date(tomorrow.getTime() + taipeiOffset);
  
  const year = taipeiTime.getUTCFullYear();
  const month = String(taipeiTime.getUTCMonth() + 1).padStart(2, '0');
  const day = String(taipeiTime.getUTCDate()).padStart(2, '0');
  const hours = String(taipeiTime.getUTCHours()).padStart(2, '0');
  
  const dateInput = document.getElementById('postpone-new-due-date');
  if (dateInput) {
    dateInput.min = `${year}-${month}-${day}T${hours}:00`;
    dateInput.value = `${year}-${month}-${day}T${hours}:00`;
  }
  
  modal.style.display = 'flex';
}

// 關閉延後歸還 Modal
function closePostponeModal() {
  const modal = document.getElementById('postpone-modal');
  if (modal) modal.style.display = 'none';
}

// 處理延後歸還表單提交 - 發送申請給 Keeper
async function handlePostponeSubmit(e) {
  e.preventDefault();
  
  const fixNo = document.getElementById('postpone-fix-no').value;
  const newDueDate = document.getElementById('postpone-new-due-date').value;
  
  if (!newDueDate) {
    alert('請選擇新的預計歸還時間');
    return;
  }
  
  const submitBtn = e.target.querySelector('button[type="submit"]');
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.textContent = '🔄 發送申請中...';
  }
  
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'requestPostpone');
    url.searchParams.append('fix_no', fixNo);
    url.searchParams.append('new_due_date', newDueDate);
    
    console.log('延後申請 URL:', url.toString());
    
    const result = await gasGetJson(url.toString());
    console.log('延後申請結果:', result);

    if (result.success) {
      alert('✅ 續借申請已送出！\n\n系統已寄信通知 Keeper 審核\n請留意您的電子郵件以接收審核結果。');
      closePostponeModal();
    } else {
      throw new Error(result.error || '申請失敗');
    }
  } catch (err) {
    if (err.isGasGlitch) {
      alert('✅ 續借申請已送出，但伺服器回應不穩定、未能確認。\n請留意 Keeper 是否收到審核信；若沒有，稍後再送一次即可。');
      closePostponeModal();
      return;
    }
    console.error('延後申請失敗:', err);
    alert('❌ 續借申請失敗：' + err.message);
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.textContent = '✅ 確認續借';
    }
  }
}

// 借用設備
async function submitBorrow(formData) {
  try {
    // 使用 GET 請求避免 CORS preflight 問題
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'borrow');
    url.searchParams.append('fix_no', formData.fix_no);
    url.searchParams.append('borrower', formData.borrower);
    url.searchParams.append('borrower_email', formData.borrower_email || '');
    url.searchParams.append('dt_borrow', formData.dt_borrow);
    url.searchParams.append('dt_due', formData.dt_due);

    console.log('借用請求網址:', url.toString());

    const result = await gasGetJson(url.toString());

    if (result.success || result.status === 'success' || (!result.error && result.message)) {
      return { success: true, message: '✅ 借用成功！已通知保管人' };
    } else {
      throw new Error(result.error || '借用失敗');
    }
  } catch (err) {
    if (err.isGasGlitch) {
      return { success: true, message: '✅ 借用已送出，但伺服器回應不穩定、未能確認結果。\n若列表未更新，請稍後重新整理。' };
    }
    console.error('借用失敗:', err);
    return { success: false, message: `❌ ${err.message}` };
  }
}

// 訪客借用請求（需要 Keeper 審核）
async function requestBorrow(formData) {
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'requestBorrow');
    url.searchParams.append('fix_no', formData.fix_no);
    url.searchParams.append('borrower', formData.borrower);
    url.searchParams.append('borrower_email', formData.borrower_email);
    url.searchParams.append('dt_borrow', formData.dt_borrow);
    url.searchParams.append('dt_due', formData.dt_due);

    console.log('借用請求網址:', url.toString());

    const result = await gasGetJson(url.toString());

    if (result.success || result.status === 'success' || (!result.error && result.message)) {
      return { success: true, message: '📧 借用申請已送出！\n\n系統已寄信通知保管人（Keeper）審核\n請留意您的電子郵件以接收審核結果。' };
    } else {
      throw new Error(result.error || '借用請求失敗');
    }
  } catch (err) {
    if (err.isGasGlitch) {
      return { success: true, message: '📧 借用申請已送出，但伺服器回應不穩定、未能確認。\n請留意 Keeper 是否收到審核信；若沒有，稍後再送一次即可。' };
    }
    console.error('借用請求失敗:', err);
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

    const text = await res.text();
    let result;
    try {
      result = JSON.parse(text);
    } catch (parseErr) {
      // GAS 偶爾會在轉址時回傳 HTML 錯誤頁，但伺服器其實通常已執行成功（信也寄出了）。
      // 這種情況視為「已送出但未能確認」，給友善提示並照常刷新列表，不要跳出嚇人的 JSON 解析錯誤。
      console.warn('歸還回應非 JSON（GAS 轉址不穩），視為已送出:', text.slice(0, 120));
      return { success: true, message: '📧 歸還通知應該已送出，但伺服器回應不穩定，系統無法確認結果。\n\n請稍後重新整理查看狀態；若設備仍顯示「借用中」，再按一次歸還即可（不會重複寄信）。' };
    }

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

    const result = await gasGetJson(url.toString());

    if (result.success || result.status === 'success' || (!result.error && result.message)) {
      alert('✅ 歸還已確認！設備狀態已更新。');
      searchEquipment();  // 重新整理列表
    } else {
      throw new Error(result.error || '確認失敗');
    }
  } catch (err) {
    if (err.isGasGlitch) {
      alert('✅ 已送出，但伺服器回應不穩定、未能確認結果。\n請稍後重新整理查看狀態。');
      searchEquipment();
      return;
    }
    console.error('確認歸還失敗:', err);
    alert(`❌ 確認失敗：${err.message}`);
  }
}

// =============================================
// 登記功能
// =============================================

// 取得登入 token（管理員專屬動作需附帶，供後端驗證身分）
function getAuthToken() {
  try {
    return (JSON.parse(localStorage.getItem('mt_user') || '{}').token) || '';
  } catch (e) {
    return '';
  }
}

// 若錯誤訊息為「未授權」（token 逾期/失效），提示並導回登入頁；有處理回傳 true
function handleAuthExpiry(msg) {
  if (msg && msg.indexOf('未授權') !== -1) {
    alert('登入已逾期，請重新登入');
    localStorage.removeItem('mt_user');
    window.location.href = './login.html';
    return true;
  }
  return false;
}

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
    url.searchParams.append('token', getAuthToken());

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
  console.log('borrow-form 元素:', borrowForm);
  
  if (borrowForm) {
    borrowForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      alert('表單提交了！');
      
      console.log('=== 借用表單提交 ===');
      console.log('借用人姓名:', document.getElementById('borrow-name')?.value);
      console.log('借用人email:', document.getElementById('borrow-email')?.value);
      console.log('預計歸還:', document.getElementById('borrow-due-date')?.value);

      const fixNo = document.getElementById('borrow-fix-no').value;
      const borrower = document.getElementById('borrow-name').value;
      const dtDue = document.getElementById('borrow-due-date').value;
      const borrowerEmailInput = document.getElementById('borrow-email');
      const borrowerEmail = borrowerEmailInput?.value?.trim() || '';
      
      // 檢查必填欄位
      console.log('檢查必填欄位 - borrower:', borrower, 'dtDue:', dtDue);
      if (!borrower) {
        console.log('借用人姓名為空，顯示警告');
        alert('請填寫借用人姓名');
        return;
      }
      if (!dtDue) {
        console.log('預計歸還日期為空，顯示警告');
        alert('請選擇預計歸還日期');
        return;
      }
      
      // 管理員借用時，如果沒有填 email，使用登入時的 email
      let finalEmail = borrowerEmail;
      if (!isGuest && !borrowerEmail && user.email) {
        finalEmail = user.email;
        console.log('管理員借用，使用登入 email:', finalEmail);
      }
      // 今天日期時間（台北時間），強制整點
      const now = new Date();
      const taipeiTime = new Date(now.getTime() + (8 * 60 * 60 * 1000));
      taipeiTime.setMinutes(0, 0, 0); // 強制整點
      const dtBorrow = taipeiTime.toISOString().slice(0, 16); // yyyy-MM-ddTHH:mm
      
      // 檢查是否為訪客（需要借用審核）
      const user = JSON.parse(localStorage.getItem('mt_user') || '{}');
      const isGuest = user.role !== 'admin';

      const submitBtn = e.target.querySelector('button[type="submit"]');
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.textContent = '🔄 處理中...';
      }

      let result;
      if (isGuest) {
        // 訪客需要發送借用請求，等待 Keeper 審核
        result = await requestBorrow({ 
          fix_no: fixNo, 
          borrower: borrower, 
          borrower_email: finalEmail,
          dt_borrow: dtBorrow, 
          dt_due: dtDue 
        });
      } else {
        // 管理員直接借用（也需要借用人 email）
        result = await submitBorrow({ 
          fix_no: fixNo, 
          borrower: borrower, 
          borrower_email: finalEmail,
          dt_borrow: dtBorrow, 
          dt_due: dtDue 
        });
      }

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
      
      // 強制使用登入者的姓名作為 keeper（防止竄改）
      const user = JSON.parse(localStorage.getItem('mt_user'));
      if (user && user.name) {
        data.keeper = user.name;
      }

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

      if (handleAuthExpiry(result.message)) return;
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

  // 頁面載入時：先從伺服器抓最新頭像快取，讓列表第一次渲染就帶頭像；
  // 即使抓頭像失敗也照常查詢設備
  preloadAvatars().finally(() => {
    searchEquipment();
    // 設備列表顯示後，於背景預先載入其他分頁，讓切換時能立即顯示
    setTimeout(preloadOtherTabs, 400);
  });
  
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
    listEl.innerHTML = loadingHtml();
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

  // 先按設備編號分組，每個設備下再按借用週期分組
  const deviceGroups = {};
  
  // 先按時間排序（舊的在前，新的在後）
  const sortedHistory = [...history].sort((a, b) => {
    return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
  });
  
  sortedHistory.forEach(record => {
    // 跳過 borrow_pending 記錄
    if (record.action === 'borrow_pending') {
      return;
    }
    
    const fixNo = record.fix_no || '無編號';
    const borrower = record.borrower || '未知';
    
    if (!deviceGroups[fixNo]) {
      deviceGroups[fixNo] = {
        fix_no: fixNo,
        device_name: record.device_name,
        cycles: []  // 每個設備有多個借用週期
      };
    }
    
    // borrow 動作表示新的借用週期開始
    if (record.action === 'borrow') {
      const newCycle = {
        borrower: borrower,
        keeper: record.keeper,
        dt_borrow: record.dt_borrow || '',
        dt_due: record.dt_due || '',
        dt_return: '',
        return_confirmed: false,
        records: [record],
        lastTimestamp: record.timestamp || ''
      };
      deviceGroups[fixNo].cycles.push(newCycle);
    } else {
      // return/confirm 歸到最新的未完成週期
      const cycles = deviceGroups[fixNo].cycles;
      const targetCycle = cycles.reverse().find(cycle => {
        return cycle.borrower === borrower && !cycle.return_confirmed;
      });
      
      if (targetCycle) {
        // 找到現有週期，更新它
        targetCycle.records.push(record);
        
        const recordTime = new Date(record.timestamp || 0).getTime();
        const currentLastTime = new Date(targetCycle.lastTimestamp || 0).getTime();
        if (recordTime > currentLastTime) {
          targetCycle.lastTimestamp = record.timestamp;
        }
        
        if (record.action === 'return') {
          targetCycle.dt_return = record.dt_return || '';
          targetCycle.return_confirmed = false;
        } else if (record.action === 'confirm' || record.action === 'confirmed') {
          targetCycle.dt_return = record.dt_return || '';
          targetCycle.return_confirmed = true;
        } else if (record.action === 'postpone' || record.action === 'postpone_approved') {
          // 延後申請核准後，更新預計歸還時間
          console.log(`[歷史] 收到 postpone_approved 記錄, fix_no=${record.fix_no}, borrower=${borrower}, dt_due=${record.dt_due}`);
          if (record.dt_due) {
            targetCycle.dt_due = record.dt_due;
            console.log(`[歷史] 更新 dt_due 為: ${record.dt_due}`);
          } else {
            console.log(`[歷史] dt_due 是空的，無法更新`);
          }
        }
      } else {
        // 沒有找到現有週期（可能只有 return/confirm，沒有 borrow），建立新週期
        const newCycle = {
          borrower: borrower,
          keeper: record.keeper,
          dt_borrow: record.dt_borrow || '',
          dt_due: record.dt_due || '',
          dt_return: record.dt_return || '',
          return_confirmed: record.action === 'confirm' || record.action === 'confirmed',
          records: [record],
          lastTimestamp: record.timestamp || ''
        };
        deviceGroups[fixNo].cycles.push(newCycle);
      }
    }
  });

  let html = '';
  
  
  // 將設備組按最新時間戳排序（取每個設備的最新借用週期）
  const sortedDeviceKeys = Object.keys(deviceGroups).sort((a, b) => {
    const deviceA = deviceGroups[a];
    const deviceB = deviceGroups[b];
    const timeA = deviceA.cycles.length > 0 ? new Date(deviceA.cycles[deviceA.cycles.length - 1].lastTimestamp || 0).getTime() : 0;
    const timeB = deviceB.cycles.length > 0 ? new Date(deviceB.cycles[deviceB.cycles.length - 1].lastTimestamp || 0).getTime() : 0;
    return sortOrder === 'newest' ? (timeB - timeA) : (timeA - timeB);
  });
  
  sortedDeviceKeys.forEach((fixNo, deviceIndex) => {
    const device = deviceGroups[fixNo];
    const deviceExpanded = deviceIndex === 0; // 第一個設備預設展開
    
    html += `
      <div class="history-device-group" style="margin-bottom:20px;">
        <div class="history-device-header" onclick="toggleHistoryDevice(this)" style="cursor:pointer;user-select:none;background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);color:white;padding:12px 15px;border-radius:6px;margin-bottom:10px;font-weight:bold;font-size:1.1em;display:flex;align-items:center;">
          <span class="device-arrow" style="display:inline-block;width:12px;margin-right:8px;transition:transform 0.2s;${deviceExpanded ? 'transform:rotate(90deg)' : ''}">▶</span>
          <span>📦 ${escapeHtml(device.fix_no)} - ${escapeHtml(device.device_name || '未知設備')}</span>
        </div>
        <div class="history-device-content" style="${deviceExpanded ? 'display:block;' : 'display:none;'}">
    `;
    
    // 每個借用週期顯示一筆，按時間排序
    const sortedCycles = [...device.cycles].sort((a, b) => {
      const timeA = new Date(a.lastTimestamp || 0).getTime();
      const timeB = new Date(b.lastTimestamp || 0).getTime();
      return sortOrder === 'newest' ? (timeB - timeA) : (timeA - timeB);
    });
    
    sortedCycles.forEach((cycle, cycleIndex) => {
      const hasConfirm = cycle.return_confirmed;
      const hasReturn = !hasConfirm && cycle.dt_return && cycle.dt_return !== '';
      const isTransfer = cycle.records && cycle.records.some(r => r.action === 'transfer');
      // 最新的預設展開（根據排序方向）
      const isExpanded = cycleIndex === 0;
      
      // 判斷狀態（transfer 最優先）
      let statusIcon, statusText;
      if (isTransfer) {
        statusIcon = '🔄';
        statusText = '轉讓';
      } else if (hasConfirm) {
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
            <span style="font-weight:bold;font-size:0.95em;">---> ${getAvatarHtml(cycle.borrower, 24)} ${escapeHtml(cycle.borrower)} ${statusIcon} ${statusText}</span>
          </div>
          <div class="history-borrow-detail" style="${isExpanded ? 'display:block;' : 'display:none;'}margin-left:20px;margin-top:8px;padding:10px;background:#fff;border-radius:6px;">
            <div style="font-size:0.9em;color:#666;line-height:1.8;">
              <div>借用日期：${formatDateTime(cycle.dt_borrow) || '－'}</div>
              <div>預計歸還：${formatDateTime(cycle.dt_due) || '－'}</div>
              ${cycle.dt_return ? `
                <div>${cycle.return_confirmed ? '歸還完成' : '歸還日期'}：${formatDateTime(cycle.dt_return)}${cycle.return_confirmed ? '' : '（待確認）'}</div>
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

// 「歷史紀錄」「我的設備」只在首次切換時載入一次；之後切分頁不重抓，
// 需重新整理頁面才會再載入（旗標隨頁面載入歸零）
let historyTabLoaded = false;
let myEquipTabLoaded = false;
let settingsTabLoaded = false;

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
      // 只在第一次切到歷史頁時載入；之後切換不重抓
      if (!historyTabLoaded) {
        searchHistory();
        historyTabLoaded = true;
      }

      // 綁定排序選單事件（只在切換到歷史頁時綁定一次）
      const historySortSelect = document.getElementById('history-sort');
      if (historySortSelect && !historySortSelect._hasEventListener) {
        historySortSelect.addEventListener('change', () => {
          searchHistory();  // 排序變更時自動重新搜尋
        });
        historySortSelect._hasEventListener = true;  // 標記已綁定，避免重複
      }
    }
    // 如果切換到「我的設備」分頁，只在第一次切換時載入；之後不重抓
    if (tab === 'my-equipment') {
      if (!myEquipTabLoaded) {
        loadMyEquipment();
        myEquipTabLoaded = true;
      }
    }
    // 如果切換到測試站列表分頁，每次切換都重新載入以取得最新狀態
    if (tab === 'test-station') {
      loadTestStations();
    }
    // 如果切換到個人設定分頁，載入頭像列表
    if (tab === 'settings') {
      // 只在第一次切到個人設定時載入頭像列表；之後不重抓
      if (!settingsTabLoaded) {
        loadAvatarList();
        settingsTabLoaded = true;
      }
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

// 從伺服器抓最新頭像並更新本地快取（頁面載入時呼叫，確保列表能正確顯示頭像）
async function preloadAvatars() {
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'getAvatarList');
    const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    const data = await res.json();
    const avatars = Array.isArray(data) ? data : (data.data || data.result || []);
    avatars.forEach(item => {
      if (item.name && item.avatar_url) {
        avatarCache[item.name] = item.avatar_url;
      }
    });
    saveAvatarCache();
  } catch (err) {
    console.log('預載頭像失敗:', err);
  }
}

// 背景預先載入其他分頁的資料，讓切換分頁時能立即顯示（不必等點選才載入）
function preloadOtherTabs() {
  const user = JSON.parse(localStorage.getItem('mt_user') || '{}');

  // 歷史紀錄（所有使用者皆可見）
  if (!historyTabLoaded) {
    searchHistory();
    historyTabLoaded = true;
  }

  // 測試站列表（所有使用者皆可見）—— 進站即背景預載，切換時可立即顯示
  loadTestStations();

  // 我的設備、個人設定（管理員專屬分頁）
  if (user.role === 'admin') {
    if (!myEquipTabLoaded) {
      loadMyEquipment();
      myEquipTabLoaded = true;
    }
    if (!settingsTabLoaded) {
      loadAvatarList();
      settingsTabLoaded = true;
    }
  }
}

/**
 * 取得頭像 HTML（圖片或預設 emoji）
 */
function getAvatarHtml(name, size = 55) {  // 圖片預設 55px
  if (!name) return `<span style="font-size:30px;">👤</span>`;  // Emoji 固定 30px
  
  const url = avatarCache[name];
  if (url) {
    return `<img src="${escapeHtml(url)}" style="width:${size}px;height:${size}px;border-radius:50%;object-fit:cover;vertical-align:middle;" onerror="this.style.display='none';this.nextElementSibling.style.display='inline';"><span style="font-size:30px;display:none;">👤</span>`;
  }
  return `<span style="font-size:30px;">👤</span>`;
}

/**
 * 載入所有頭像列表（從 Google Sheet）
 */
async function loadAvatarList() {
  const listEl = document.getElementById('avatar-list');
  if (!listEl) return;
  
  // 自動填入登入者姓名
  const user = JSON.parse(localStorage.getItem('mt_user') || '{}');
  if (user && user.name) {
    const nameInput = document.getElementById('avatar-name');
    if (nameInput) {
      nameInput.value = user.name;
    }
  }
  
  listEl.innerHTML = loadingHtml();
  
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
          <img src="${escapeHtml(url)}" style="width:60px;height:60px;border-radius:50%;object-fit:cover;border:2px solid #ddd;">
          <div style="font-size:0.85em;margin-top:5px;">${escapeHtml(name)}</div>
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

// =============================================
// 我的設備功能（管理員專屬）
// =============================================

/**
 * 載入管理員自己的設備列表
 */
async function loadMyEquipment() {
  console.log('loadMyEquipment 開始');
  const user = JSON.parse(localStorage.getItem('mt_user'));
  console.log('使用者資料:', user);
  
  if (!user || user.role !== 'admin') {
    console.log('不是管理員或未登入');
    document.getElementById('my-equipment-list').innerHTML = 
      '<p style="text-align:center;color:#c00;padding:40px;">❌ 只有管理員可以查看此頁面</p>';
    return;
  }
  
  const listEl = document.getElementById('my-equipment-list');
  listEl.innerHTML = loadingHtml();
  
  try {
    // 查詢所有設備（從兩個工作表）
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'query');
    console.log('查詢 URL:', url.toString());
    
    const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    console.log('GAS 回應狀態:', res.status);
    
    if (!res.ok) {
      throw new Error('HTTP ' + res.status);
    }
    
    const data = await res.json();
    console.log('GAS 回應資料:', data);
    
    if (data.error) {
      throw new Error(data.error);
    }
    
    const allEquipment = Array.isArray(data) ? data : (data.data || []);
    console.log('總設備數量:', allEquipment.length);
    console.log('登入者姓名:', user.name);
    
    // 只顯示 keeper = 登入者的設備
    const myEquipment = allEquipment.filter(eq => eq.keeper === user.name);
    console.log('我的設備數量:', myEquipment.length);
    
    if (myEquipment.length === 0) {
      listEl.innerHTML = `
        <div style="text-align:center;padding:40px;color:#666;">
          <p style="font-size:1.2em;">📭 您還沒有保管任何設備</p>
          <p style="margin-top:10px;">點擊「設備登記」新增設備</p>
        </div>
      `;
      return;
    }
    
    // 顯示設備列表
    let html = '<div style="display:grid;gap:15px;">';
    console.log('所有設備狀態:', allEquipment.map(eq => eq.status));
    myEquipment.forEach((eq, index) => {
      const isBorrowed = eq.status === 'borrowed' || eq.status === '借用中' || eq.status === '已借出' || eq.status === '使用中';
      const isReturnPending = eq.status === 'return_pending';
      
      html += `
        <div style="background:#fff;border:1px solid #ddd;border-radius:8px;padding:15px;">
          <div style="display:flex;justify-content:space-between;align-items:start;">
            <div>
              <strong style="font-size:1.1em;">${escapeHtml(eq.device_name || '未知設備')}</strong>
              <div style="color:#666;font-size:0.85em;margin-top:5px;">
                📋 編號：${escapeHtml(eq.fix_no || '-')}
              </div>
              <div style="color:#666;font-size:0.85em;">
                📦 類型：${escapeHtml(eq.fix_type || '-')}
              </div>
            </div>
            <div style="text-align:right;">
              ${isBorrowed ? '<span style="background:#ffc107;padding:4px 8px;border-radius:4px;font-size:0.85em;">📤 已借出</span>' : ''}
              ${isReturnPending ? '<span style="background:#17a2b8;padding:4px 8px;border-radius:4px;font-size:0.85em;color:white;">⏳ 歸還中</span>' : ''}
              ${eq.status === 'available' ? '<span style="background:#28a745;padding:4px 8px;border-radius:4px;font-size:0.85em;color:white;">✓ 可借用</span>' : ''}
            </div>
          </div>
          <div style="margin-top:15px;display:flex;gap:10px;">
            ${(isBorrowed || isReturnPending) ? 
              '<button disabled style="padding:8px 15px;background:#ccc;color:#888;border:none;border-radius:6px;cursor:not-allowed;" title="使用中無法修改">✏️ 修改</button><button disabled style="padding:8px 15px;background:#ccc;color:#888;border:none;border-radius:6px;cursor:not-allowed;" title="使用中無法刪除">🗑️ 刪除</button><button disabled style="padding:8px 15px;background:#ccc;color:#888;border:none;border-radius:6px;cursor:not-allowed;" title="使用中無法轉讓">🔄 轉讓</button>' :
              '<button onclick="openEditEquipmentModal(\'' + eq.fix_no + '\', \'' + encodeURIComponent(eq.device_name || '') + '\', \'' + encodeURIComponent(eq.fix_type || '') + '\', \'' + (eq.qty_asset || '1') + '\')" style="padding:8px 15px;background:#667eea;color:white;border:none;border-radius:6px;cursor:pointer;">✏️ 修改</button><button onclick="confirmDeleteEquipment(\'' + eq.fix_no + '\', \'' + eq.device_name + '\')" style="padding:8px 15px;background:#dc3545;color:white;border:none;border-radius:6px;cursor:pointer;">🗑️ 刪除</button><button onclick="openTransferModal(\'' + eq.fix_no + '\', \'' + encodeURIComponent(eq.device_name || '') + '\', \'' + encodeURIComponent(eq.keeper || '') + '\')" style="padding:8px 15px;background:#17a2b8;color:white;border:none;border-radius:6px;cursor:pointer;">🔄 轉讓</button>'
            }
          </div>
        </div>
      `;
    });
    html += '</div>';
    
    listEl.innerHTML = html;
  } catch (err) {
    console.error('載入我的設備失敗:', err);
    listEl.innerHTML = `<p style="text-align:center;color:#c00;padding:40px;">❌ 載入失敗：${err.message}</p>`;
  }
}

/**
 * 開啟修改設備 Modal
 */
function openEditEquipmentModal(fixNo, deviceName, fixType, qtyAsset) {
  document.getElementById('edit-fix-no').value = fixNo;
  document.getElementById('edit-fix-no-display').value = fixNo || '';
  document.getElementById('edit-device-name').value = decodeURIComponent(deviceName || '');
  document.getElementById('edit-fix-type').value = decodeURIComponent(fixType || '儀器設備');
  document.getElementById('edit-qty-asset').value = qtyAsset || '1';
  document.getElementById('edit-equipment-modal').style.display = 'flex';
}

function closeEditEquipmentModal() {
  document.getElementById('edit-equipment-modal').style.display = 'none';
}

/**
 * 更新設備
 */
async function updateEquipment(fixNo, updateData) {
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'updateEquipment');
    url.searchParams.append('fix_no', fixNo);
    url.searchParams.append('device_name', updateData.device_name || '');
    url.searchParams.append('fix_type', updateData.fix_type || '');
    url.searchParams.append('token', getAuthToken());

    const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    const data = await res.json();
    
    if (data.success) {
      alert('✅ 設備已更新');
      loadMyEquipment();  // 重新載入列表
    } else {
      if (handleAuthExpiry(data.error)) return;
      alert('❌ 更新失敗：' + (data.error || '未知錯誤'));
    }
  } catch (err) {
    alert('❌ 更新失敗：' + err.message);
  }
}

/**
 * 確認刪除設備
 */
function confirmDeleteEquipment(fixNo, deviceName) {
  if (!confirm('確定要刪除設備「' + deviceName + '」嗎？\n此操作無法撤銷！')) {
    return;
  }
  deleteEquipment(fixNo);
}

/**
 * 刪除設備
 */
async function deleteEquipment(fixNo) {
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'deleteEquipment');
    url.searchParams.append('fix_no', fixNo);
    url.searchParams.append('token', getAuthToken());

    const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    const data = await res.json();
    
    if (data.success) {
      alert('✅ 設備已刪除');
      loadMyEquipment();  // 重新載入列表
    } else {
      if (handleAuthExpiry(data.error)) return;
      alert('❌ 刪除失敗：' + (data.error || '未知錯誤'));
    }
  } catch (err) {
    alert('❌ 刪除失敗：' + err.message);
  }
}

// 修改設備表單提交
document.getElementById('edit-equipment-form').addEventListener('submit', async function(e) {
  e.preventDefault();
  console.log('表單提交了');
  // originalFixNo 是隱藏欄位，用來找資料列
  // newFixNo 是顯示欄位，是使用者輸入的新編號
  const originalFixNo = document.getElementById('edit-fix-no').value;
  const newFixNo = document.getElementById('edit-fix-no-display').value;
  const deviceName = document.getElementById('edit-device-name').value;
  const fixType = document.getElementById('edit-fix-type').value;
  const qtyAsset = document.getElementById('edit-qty-asset').value;
  
  if (!originalFixNo || !deviceName) {
    alert('請填寫設備名稱');
    return;
  }
  
    try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'updateEquipment');
    url.searchParams.append('fix_no', originalFixNo);
    url.searchParams.append('new_fix_no', newFixNo);
    url.searchParams.append('device_name', deviceName);
    url.searchParams.append('fix_type', fixType);
    url.searchParams.append('qty_asset', qtyAsset || '1');
    url.searchParams.append('token', getAuthToken());
    
    console.log('更新設備 URL:', url.toString());
    
    const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    console.log('更新設備 HTTP 狀態:', res.status);
    const data = await res.json();
    console.log('更新設備回應:', data);
    
    if (data.success) {
      alert('✅ 設備已更新');
      closeEditEquipmentModal();
      loadMyEquipment();
    } else {
      if (handleAuthExpiry(data.error)) return;
      alert('❌ 更新失敗：' + (data.error || '未知錯誤'));
    }
  } catch (err) {
    console.error('更新設備錯誤:', err);
    alert('❌ 更新失敗：' + err.message);
  }
});

// =============================================
// 設備轉讓功能
// =============================================

/**
 * 開啟轉讓設備 Modal
 */
function openTransferModal(fixNo, deviceName, keeper) {
  // 動態建立 Modal（如果還沒有）
  let modal = document.getElementById('transfer-modal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'transfer-modal';
    modal.style.cssText = 'display:none;position:fixed;z-index:1000;left:0;top:0;width:100%;height:100%;overflow-y:auto;padding:20px;background-color:rgba(0,0,0,0.5);';
    modal.innerHTML = `
      <div style="background-color:#fefefe;margin:auto;padding:20px;border:1px solid #888;width:90%;max-width:500px;border-radius:12px;position:relative;">
        <span onclick="closeTransferModal()" style="color:#aaa;float:right;font-size:28px;font-weight:bold;cursor:pointer;">&times;</span>
        <h2 style="color:#17a2b8;margin-bottom:20px;">🔄 轉讓設備</h2>
        <div id="transfer-info" style="background:#f8f9fa;padding:15px;border-radius:8px;margin-bottom:20px;"></div>
        <form id="transfer-form">
          <input type="hidden" id="transfer-fix-no">
          <div style="margin-bottom:15px;">
            <label style="display:block;margin-bottom:5px;font-weight:bold;">選擇接收 Keeper：</label>
            <select id="transfer-to-keeper" required style="width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;font-size:1em;">
              <option value="">載入中...</option>
            </select>
          </div>
          <div style="text-align:center;margin-top:20px;">
            <button type="submit" class="btn" style="background:#17a2b8;color:#fff;padding:12px 24px;border:none;border-radius:8px;cursor:pointer;font-size:1em;">✅ 送出轉讓申請</button>
            <button type="button" onclick="closeTransferModal()" style="background:#6c757d;color:#fff;padding:12px 24px;border:none;border-radius:8px;cursor:pointer;font-size:1em;margin-left:10px;">取消</button>
          </div>
        </form>
      </div>
    `;
    document.body.appendChild(modal);
    
    // 綁定表單提交事件
    document.getElementById('transfer-form').addEventListener('submit', handleTransferSubmit);
  }
  
  // 設定設備資訊
  document.getElementById('transfer-fix-no').value = fixNo;
  document.getElementById('transfer-info').innerHTML = `
    <strong>設備編號：</strong>${escapeHtml(fixNo)}<br>
    <strong>設備名稱：</strong>${escapeHtml(decodeURIComponent(deviceName || ''))}<br>
    <strong>目前保管人：</strong>${escapeHtml(decodeURIComponent(keeper || ''))}
  `;
  
  // 載入 Keeper 清單
  loadKeeperListForTransfer(fixNo);
  
  modal.style.display = 'flex';
}

/**
 * 載入 Keeper 清單（排除自己）
 */
async function loadKeeperListForTransfer(currentFixNo) {
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'getKeeperList');
    
    const res = await fetch(url.toString());
    const data = await res.json();
    
    const select = document.getElementById('transfer-to-keeper');
    const user = JSON.parse(localStorage.getItem('mt_user'));
    
    if (data.success && data.keepers) {
      // 過濾掉自己
      const keepers = data.keepers.filter(k => k.name !== user.name);
      
      if (keepers.length === 0) {
        select.innerHTML = '<option value="">沒有其他 Keeper</option>';
        return;
      }
      
      select.innerHTML = keepers.map(k =>
        `<option value="${escapeHtml(k.name)}">${escapeHtml(k.name)}</option>`
      ).join('');
    } else {
      select.innerHTML = '<option value="">載入失敗</option>';
    }
  } catch (err) {
    console.error('載入 Keeper 清單失敗:', err);
    document.getElementById('transfer-to-keeper').innerHTML = '<option value="">載入失敗</option>';
  }
}

/**
 * 關閉轉讓 Modal
 */
function closeTransferModal() {
  const modal = document.getElementById('transfer-modal');
  if (modal) modal.style.display = 'none';
}

/**
 * 處理轉讓表單提交
 */
async function handleTransferSubmit(e) {
  e.preventDefault();
  
  const fixNo = document.getElementById('transfer-fix-no').value;
  const toKeeper = document.getElementById('transfer-to-keeper').value;
  
  if (!toKeeper) {
    alert('請選擇要轉讓給哪位 Keeper');
    return;
  }
  
  const submitBtn = e.target.querySelector('button[type="submit"]');
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.textContent = '🔄 處理中...';
  }
  
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'requestTransfer');
    url.searchParams.append('fix_no', fixNo);
    url.searchParams.append('to_keeper', toKeeper);
    url.searchParams.append('token', getAuthToken());

    const res = await fetch(url.toString());
    const result = await res.json();
    
    if (result.success) {
      alert('✅ 轉讓申請已送出！\n\n系統已寄信通知 ' + toKeeper + ' 審核');
      closeTransferModal();
      loadMyEquipment();
    } else {
      throw new Error(result.error || '申請失敗');
    }
  } catch (err) {
    if (handleAuthExpiry(err.message)) return;
    alert('❌ 轉讓失敗：' + err.message);
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.textContent = '✅ 送出轉讓申請';
    }
  }
}

// =============================================
// 部門儀器借用功能（任何人可用）
// =============================================

/**
 * 處理部門儀器借用提交
 */
async function handleDeptBorrowSubmit() {
  console.log('=== handleDeptBorrowSubmit 被呼叫 ===');
  
  const deviceName = document.getElementById('dept-borrow-device').value.trim();
  const borrower = document.getElementById('dept-borrow-name').value.trim();
  const borrowerEmail = document.getElementById('dept-borrow-email').value.trim();
  const dtDue = document.getElementById('dept-borrow-due-date').value;
  
  console.log('設備名稱:', deviceName);
  console.log('借用人:', borrower);
  console.log('Email:', borrowerEmail);
  console.log('預計歸還:', dtDue);
  
  // 檢查必填欄位
  if (!deviceName) {
    alert('請填寫設備名稱');
    return;
  }
  if (!borrower) {
    alert('請填寫借用人姓名');
    return;
  }
  if (!borrowerEmail) {
    alert('請填寫電子郵件');
    return;
  }
  if (!dtDue) {
    alert('請選擇預計歸還日期時間');
    return;
  }
  
  // 借用日期：四捨五入到最接近的整點
  const now = new Date();
  const hours = now.getHours();
  const minutes = now.getMinutes();
  let borrowHour = hours;
  if (minutes >= 30) {
    borrowHour = hours + 1; // >=30分鐘，小時+1
  }
  // 處理跨日
  let borrowDay = now.getDate();
  if (borrowHour >= 24) {
    borrowHour = 0;
    borrowDay++;
  }
  // 格式化為 yyyy-MM-ddTHH:mm
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(borrowDay).padStart(2, '0');
  const hour = String(borrowHour).padStart(2, '0');
  const dtBorrow = `${year}-${month}-${day}T${hour}:00`;
  
  const btn = document.querySelector('#dept-borrow-form button');
  if (btn) {
    btn.disabled = true;
    btn.textContent = '🔄 處理中...';
  }
  
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'deptBorrow');
    url.searchParams.append('device_name', deviceName);
    url.searchParams.append('borrower', borrower);
    url.searchParams.append('borrower_email', borrowerEmail);
    url.searchParams.append('dt_borrow', dtBorrow);
    url.searchParams.append('dt_due', dtDue);
    
    console.log('部門儀器借用 URL:', url.toString());
    
    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });
    
    const text = await res.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch (parseErr) {
      // GAS 轉址不穩回傳 HTML，但伺服器通常已借用成功（借用人會收到確認信）
      console.warn('部門儀器借用回應非 JSON（GAS 不穩），視為已送出:', text.slice(0, 120));
      alert('✅ 借用已送出！\n\n伺服器回應不穩定，若您已收到確認信即代表借用成功。\n若下方列表未更新，請稍後重新整理。');
      document.getElementById('dept-borrow-form').reset();
      loadDeptBorrowList();
      return;
    }
    console.log('部門儀器借用回應:', data);

    if (data.success) {
      alert('✅ 借用成功！確認郵件已寄出');
      // 清空表單
      document.getElementById('dept-borrow-form').reset();
      // 重新載入列表
      loadDeptBorrowList();
    } else {
      alert('❌ 借用失敗：' + (data.error || '未知錯誤'));
    }
  } catch (err) {
    console.error('部門儀器借用錯誤:', err);
    alert('❌ 借用失敗：' + err.message);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = '✅ 確認借用';
    }
  }
}

/**
 * 載入部門儀器借用列表
 */
async function loadDeptBorrowList() {
  const listEl = document.getElementById('dept-borrow-list');
  if (!listEl) return;
  
  listEl.innerHTML = loadingHtml();
  
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'getDeptBorrowList');
    
    const res = await fetch(url.toString(), {
      method: 'GET',
      redirect: 'follow'
    });
    
    const data = await res.json();
    console.log('部門儀器列表:', data);
    
    if (!data.success || !data.items || data.items.length === 0) {
      listEl.innerHTML = '<p style="text-align:center;color:#666;padding:40px;">目前沒有借用的部門儀器</p>';
      return;
    }
    
    // 生成 HTML
    const html = data.items.map(item => {
      const isOverdue = new Date(item.dt_due) < new Date() && !item.dt_return;
      const statusClass = isOverdue ? 'style="color:#c00;"' : 'style="color:#0a0;"';
      const statusText = item.dt_return ? '✅ 已歸還' : (isOverdue ? '⏰ 已逾期' : '📤 借用中');
      
      // 根據是否已歸還，決定日期顯示內容
      let dateHtml = `
        <div style="color:#666;font-size:0.85em;margin-top:3px;">📅 借用日期：${item.dt_borrow}</div>
        <div style="color:#666;font-size:0.85em;margin-top:3px;">📅 預計歸還：${item.dt_due}</div>
      `;
      
      if (item.dt_return) {
        dateHtml += `<div style="color:#666;font-size:0.85em;margin-top:3px;">📅 歸還日期：${item.dt_return}</div>`;
      }
      
      return `
        <div class="dept-borrow-item" style="background:${item.dt_return ? '#e8f5e9' : (isOverdue ? '#ffebee' : '#f8f9fa')};border-left:4px solid ${item.dt_return ? '#4caf50' : (isOverdue ? '#f44336' : '#667eea')};border-radius:8px;padding:15px;margin-bottom:10px;">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px;">
            <div style="flex:1;min-width:200px;">
              <strong style="font-size:1.1em;">${escapeHtml(item.device_name)}</strong>
              <div style="color:#666;font-size:0.9em;margin-top:5px;">
                👤 借用人：${escapeHtml(item.borrower)}
              </div>
              ${dateHtml}
            </div>
            <div style="text-align:right;">
              <span ${statusClass} style="font-weight:bold;">${statusText}</span>
              ${!item.dt_return ? `
                <button class="btn-return-sm" onclick="handleDeptReturn('${item.id}', '${escapeHtml(item.device_name)}', '${escapeHtml(item.borrower)}')" style="margin-left:10px;padding:5px 10px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;font-size:0.85em;">
                  歸還
                </button>
              ` : ''}
            </div>
          </div>
        </div>
      `;
    }).join('');
    
    listEl.innerHTML = html;
    
  } catch (err) {
    console.error('載入部門儀器列表失敗:', err);
    listEl.innerHTML = '<p style="text-align:center;color:#c00;padding:40px;">❌ 載入失敗</p>';
  }
}

/**
 * 處理部門儀器歸還
 */
async function handleDeptReturn(id, deviceName, borrower) {
  if (!confirm(`確認要歸還「${deviceName}」嗎？`)) {
    return;
  }
  
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'deptReturn');
    url.searchParams.append('id', id);
    
    const data = await gasGetJson(url.toString());

    if (data.success) {
      alert('✅ 歸還成功！已通知管理員');
      loadDeptBorrowList();
    } else {
      alert('❌ 歸還失敗：' + (data.error || '未知錯誤'));
    }
  } catch (err) {
    if (err.isGasGlitch) {
      alert('✅ 歸還已送出，但伺服器回應不穩定、未能確認。\n請稍後重新整理查看狀態；若仍在借用中，再按一次即可。');
      loadDeptBorrowList();
      return;
    }
    console.error('歸還失敗:', err);
    alert('❌ 歸還失敗：' + err.message);
  }
}

/**
 * HTML 跳脫
 */
function escapeHtml(text) {
  if (!text) return '';
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// 頁面載入時載入部門儀器列表
document.addEventListener('DOMContentLoaded', function() {
  // 設定最小日期時間為現在，強制整點
  const today = new Date();
  const taipeiTime = new Date(today.getTime() + (8 * 60 * 60 * 1000));
  taipeiTime.setMinutes(0, 0, 0); // 強制整點
  const taipeiDateTime = taipeiTime.toISOString().slice(0, 16); // yyyy-MM-ddTHH:mm
  
  const deptDateInput = document.getElementById('dept-borrow-due-date');
  if (deptDateInput) {
    deptDateInput.min = taipeiDateTime;
    
    // 監聽變更事件，強制改為最接近的整點（四捨五入）- deptDateInput
    deptDateInput.addEventListener('change', function() {
      if (this.value) {
        // datetime-local 格式: 2026-05-20T14:46 或 2026-05-20T23:36
        const parts = this.value.split('T');
        const datePart = parts[0];
        const timePart = parts[1];
        const hourPart = timePart.substring(0, 2);
        const minPart = timePart.substring(3, 5);
        
        const hours = parseInt(hourPart, 10);
        const minutes = parseInt(minPart, 10);
        
        // 計算新的整點時間
        let newHours = hours;
        if (minutes >= 30) {
          newHours = hours + 1;
        }
        
        // 處理跨日（例如 23:30 -> 00:00 隔天）
        let newDatePart = datePart;
        if (newHours >= 24) {
          newHours = 0;
          // 手動計算隔天日期（避免 Date 物件時區問題）
          const [y, m, d] = datePart.split('-').map(Number);
          let newDay = d + 1;
          let newMonth = m;
          let newYear = y;
          
          // 檢查是否跨月（例如 1/31 -> 2/1）
          const daysInMonth = new Date(y, m, 0).getDate();
          if (newDay > daysInMonth) {
            newDay = 1;
            newMonth = m + 1;
            // 檢查是否跨年
            if (newMonth > 12) {
              newMonth = 1;
              newYear = y + 1;
            }
          }
          
          const yStr = String(newYear);
          const mStr = String(newMonth).padStart(2, '0');
          const dStr = String(newDay).padStart(2, '0');
          newDatePart = `${yStr}-${mStr}-${dStr}`;
        }
        
        // 格式化新值
        const newHourStr = String(newHours).padStart(2, '0');
        const newValue = newDatePart + 'T' + newHourStr + ':00';
        this.value = newValue;
      }
    });
  }
  
  // 載入列表（只在頁面切換時載入）
  loadDeptBorrowList();
});

// =============================================
// 測試站（座位）借用功能
// =============================================

// 取得「現在」整點後的 datetime-local 字串（台北時間），作為 min 值
function stationMinDateTime() {
  const now = new Date();
  const taipei = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  taipei.setUTCMinutes(0, 0, 0); // 強制整點
  return taipei.toISOString().slice(0, 16); // yyyy-MM-ddTHH:mm
}

// 將 datetime-local 的值強制吸附到最接近的整點（比照儀器借用規則）
// 規則：分鐘 >= 30 → 進位到下一小時 :00，否則捨去為當前小時 :00（含跨日/跨月/跨年）
function snapToHour(input) {
  if (!input || !input.value) return;
  const [datePart, timePart] = input.value.split('T');
  const hours = parseInt(timePart.substring(0, 2), 10);
  const minutes = parseInt(timePart.substring(3, 5), 10);

  let newHours = hours;
  if (minutes >= 30) {
    newHours = hours + 1;
  }

  // 處理跨日（例如 23:30 → 00:00 隔天）
  let newDatePart = datePart;
  if (newHours >= 24) {
    newHours = 0;
    const [y, m, d] = datePart.split('-').map(Number);
    let newDay = d + 1, newMonth = m, newYear = y;
    const daysInMonth = new Date(y, m, 0).getDate();
    if (newDay > daysInMonth) {
      newDay = 1;
      newMonth = m + 1;
      if (newMonth > 12) { newMonth = 1; newYear = y + 1; }
    }
    newDatePart = `${newYear}-${String(newMonth).padStart(2, '0')}-${String(newDay).padStart(2, '0')}`;
  }

  const hStr = String(newHours).padStart(2, '0');
  input.value = `${newDatePart}T${hStr}:00`;
}

// 載入測試站列表（讀取為冪等操作，遇到 GAS 轉址不穩回傳 HTML 時自動重試）
async function loadTestStations() {
  const list = document.getElementById('test-station-list');
  // 已經有卡片時（例如背景預載完成後再切換進來）就靜默更新，不再閃轉圈圈
  const hasContent = list && list.querySelector('.station-card');
  if (list && !hasContent) list.innerHTML = loadingHtml();

  const maxTries = 3;
  let lastErr = '載入失敗';

  for (let attempt = 1; attempt <= maxTries; attempt++) {
    try {
      const url = new URL(GAS_URL);
      url.searchParams.append('action', 'queryStations');
      const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
      const text = await res.text();

      let data;
      try {
        data = JSON.parse(text);
      } catch (parseErr) {
        // GAS 偶爾轉址不穩回傳 HTML（暫時性）→ 稍等後重試
        lastErr = '伺服器忙碌，請稍後再試';
        if (attempt < maxTries) { await new Promise(r => setTimeout(r, 1200)); continue; }
        break;
      }

      if (data.error) { lastErr = data.error; break; }
      const stations = Array.isArray(data) ? data : (data.data || []);
      renderTestStations(stations);
      return; // 成功
    } catch (err) {
      // 網路層錯誤也重試
      lastErr = err.message || '載入失敗';
      if (attempt < maxTries) { await new Promise(r => setTimeout(r, 1200)); continue; }
    }
  }

  console.error('載入測試站失敗:', lastErr);
  if (list) {
    list.innerHTML = `<p style="text-align:center;color:#c00;padding:40px;">❌ 載入失敗：${lastErr}<br>
      <button onclick="loadTestStations()" style="margin-top:12px;padding:8px 18px;border:none;border-radius:6px;background:#667eea;color:#fff;cursor:pointer;">🔄 重新載入</button></p>`;
  }
}

// 判斷測試站是否已逾期（現在時間超過預計歸還時間）
function isStationOverdue(dtDue) {
  if (!dtDue) return false;
  const s = dtDue.trim();
  // 純日期（以天為單位）→ 當天結束前都算有效，超過當天才算逾期
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const due = new Date(s + 'T23:59:59');
    if (isNaN(due.getTime())) return false;
    return Date.now() > due.getTime();
  }
  // 相容舊的「日期＋時間」資料
  const due = new Date(s.replace(' ', 'T'));
  if (isNaN(due.getTime())) return false;
  return Date.now() > due.getTime();
}

// 最近一次載入的測試站資料（供登記視窗判斷哪些日期已被登記）
let lastStationsData = [];

// 渲染測試站卡片
function renderTestStations(stations) {
  const list = document.getElementById('test-station-list');
  if (!list) return;

  lastStationsData = stations || [];

  if (!stations || stations.length === 0) {
    list.innerHTML = '<p style="text-align:center;color:#666;padding:40px;">目前沒有測試站</p>';
    return;
  }

  const today = stationMinDateTime().slice(0, 10); // 台北時間今天 yyyy-MM-dd
  let html = '';
  stations.forEach(s => {
    const stationName = escapeHtml(s.station || '');
    const bookings = s.bookings || [];
    const inUseToday = bookings.some(b => b.date === today);
    let bookingHtml = '';
    if (bookings.length === 0) {
      bookingHtml = '<div class="station-empty">目前無人登記</div>';
    } else {
      bookings.forEach(b => {
        const purposeHtml = b.purpose
          ? `<div class="booking-purpose">📝 ${escapeHtml(b.purpose)}</div>`
          : '';
        bookingHtml += `
          <div class="booking-row">
            <button class="booking-cancel" onclick="cancelStationBooking('${b.id}', '${b.date}')" title="取消這筆登記">✕</button>
            <div class="booking-main">
              <span class="booking-date">📅 ${escapeHtml(b.date)}</span>
              <span class="booking-name">${escapeHtml(b.booker || '')}</span>
            </div>
            ${purposeHtml}
          </div>`;
      });
    }
    html += `
      <div class="station-card${inUseToday ? ' in-use' : ''}">
        <div class="station-avatar">${inUseToday ? '🧑‍💻' : '🖥️'}</div>
        <div class="station-today ${inUseToday ? 'busy' : 'free'}">${inUseToday ? '今日使用中' : '今日空閒'}</div>
        <div class="station-title">${stationName}</div>
        <div class="booking-header">📌 已被預約日期${bookings.length ? `（${bookings.length}）` : ''}</div>
        <div class="booking-list">${bookingHtml}</div>
        <button class="btn-borrow-sm station-book-btn" onclick="openStationBookModal('${s.station}')">➕ 登記使用</button>
      </div>`;
  });
  list.innerHTML = html;
}

// 動態建立測試站用的 Modal 容器
function buildStationModal(id, innerHtml) {
  let modal = document.getElementById(id);
  if (!modal) {
    modal = document.createElement('div');
    modal.id = id;
    modal.className = 'modal';
    modal.style.cssText = 'display:none;position:fixed;z-index:1000;left:0;top:0;width:100%;height:100%;overflow-y:auto;padding:20px;background-color:rgba(0,0,0,0.5);';
    document.body.appendChild(modal);
  }
  modal.innerHTML = `
    <div style="background-color:#fefefe;margin:auto;padding:20px;border:1px solid #888;width:90%;max-width:460px;border-radius:12px;position:relative;">
      <span onclick="document.getElementById('${id}').style.display='none'" style="color:#aaa;float:right;font-size:28px;font-weight:bold;cursor:pointer;">&times;</span>
      ${innerHtml}
    </div>`;
  modal.style.display = 'block';
  return modal;
}

// 開啟測試站「登記使用」Modal（選開始～結束日期的連續區間）
function openStationBookModal(station) {
  const minDate = stationMinDateTime().slice(0, 10);
  buildStationModal('station-book-modal', `
    <h2 style="color:#667eea;margin-bottom:20px;">➕ 登記使用 ${escapeHtml(station)}</h2>
    <form onsubmit="submitStationBook(event, '${station}')">
      <div class="form-group">
        <label>登記人姓名 *</label>
        <input type="text" id="station-book-name" required placeholder="請輸入您的姓名" style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ddd;border-radius:4px;">
      </div>
      <div class="form-group" style="margin-top:12px;">
        <label>使用日期（起訖）* <span style="font-size:0.85em;color:#888;">(只用一天：結束可留空)</span></label>
        <div style="display:flex;gap:8px;align-items:center;">
          <input type="date" id="station-book-start" min="${minDate}" required style="flex:1;box-sizing:border-box;padding:10px;border:1px solid #ddd;border-radius:4px;">
          <span style="color:#888;">～</span>
          <input type="date" id="station-book-end" min="${minDate}" style="flex:1;box-sizing:border-box;padding:10px;border:1px solid #ddd;border-radius:4px;">
        </div>
      </div>
      <div class="form-group" style="margin-top:12px;">
        <label>用途 / 備註</label>
        <input type="text" id="station-book-purpose" placeholder="選填，例如：EE check、FT test" style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ddd;border-radius:4px;">
      </div>
      <div style="text-align:center;margin-top:20px;">
        <button type="submit" class="btn-primary">✅ 確認登記</button>
      </div>
    </form>
  `);
}

// 把 start~end（含）展開成 yyyy-MM-dd 陣列
function expandDateRange(start, end) {
  const out = [];
  let d = new Date(start + 'T00:00:00');
  const endD = new Date(end + 'T00:00:00');
  while (d <= endD) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    out.push(`${y}-${m}-${day}`);
    d.setDate(d.getDate() + 1);
  }
  return out;
}

// 送出測試站登記（開始～結束的連續區間）
async function submitStationBook(e, station) {
  e.preventDefault();
  const start = document.getElementById('station-book-start').value;
  let end = document.getElementById('station-book-end').value;
  const name = document.getElementById('station-book-name').value.trim();
  const purpose = document.getElementById('station-book-purpose').value.trim();

  if (!start) { alert('請選擇開始日期'); return; }
  if (!end) end = start; // 結束留空 = 只登記一天
  if (end < start) { alert('結束日期不能早於開始日期'); return; }
  if (!name) { alert('請填寫登記人姓名'); return; }

  const dates = expandDateRange(start, end);
  if (dates.length > 90) { alert('一次最多登記 90 天，請縮短區間'); return; }

  const btn = e.target.querySelector('button[type="submit"]');
  if (btn) { btn.disabled = true; btn.textContent = '🔄 處理中...'; }

  const closeAndRefresh = () => {
    const modal = document.getElementById('station-book-modal');
    if (modal) modal.style.display = 'none';
    loadTestStations();
  };

  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'bookStation');
    url.searchParams.append('station', station);
    url.searchParams.append('dates', dates.join(','));
    url.searchParams.append('booker', name);
    url.searchParams.append('purpose', purpose);
    // bookStation 是冪等的（同人同日重送不會重複），故遇 GAS 抽風可安全自動重試
    const result = await gasGetJson(url.toString(), { retries: 3 });
    if (result.success) {
      if (result.conflicts && result.conflicts.length) {
        const c = result.conflicts.map(x => `${x.date}（已被 ${x.by}）`).join('\n');
        alert(`以下日期已被登記、未成功：\n${c}\n\n其餘日期已登記完成。`);
      }
      closeAndRefresh();
    } else {
      throw new Error(result.error || '登記失敗');
    }
  } catch (err) {
    if (err.isGasGlitch) { closeAndRefresh(); return; }
    alert('❌ 登記失敗：' + err.message);
    if (btn) { btn.disabled = false; btn.textContent = '✅ 確認登記'; }
  }
}

// 取消一筆測試站登記
async function cancelStationBooking(id, date) {
  if (!confirm(`確定要取消 ${date} 的登記嗎？`)) return;
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'cancelStationBooking');
    url.searchParams.append('id', id);
    const result = await gasGetJson(url.toString());
    if (result.success) {
      loadTestStations();
    } else {
      throw new Error(result.error || '取消失敗');
    }
  } catch (err) {
    if (err.isGasGlitch) { loadTestStations(); return; }
    alert('❌ 取消失敗：' + err.message);
  }
}

// 開啟測試站借用 Modal
function openStationBorrowModal(station) {
  buildStationModal('station-borrow-modal', `
    <h2 style="color:#667eea;margin-bottom:20px;">🖥️ 借用測試站 ${station}</h2>
    <form onsubmit="submitStationBorrow(event, '${station}')">
      <div class="form-group">
        <label>借用人姓名 *</label>
        <input type="text" id="station-borrow-name" required placeholder="請輸入您的姓名" style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ddd;border-radius:4px;">
      </div>
      <div class="form-group" style="margin-top:12px;">
        <label>預計歸還日期 * <span style="font-size:0.85em;color:#888;">(以天為單位)</span></label>
        <input type="date" id="station-borrow-due" min="${stationMinDateTime().slice(0, 10)}" required style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ddd;border-radius:4px;">
      </div>
      <div style="text-align:center;margin-top:20px;">
        <button type="submit" class="btn-primary">✅ 確認借用</button>
      </div>
    </form>
  `);
}

// 送出測試站借用
async function submitStationBorrow(e, station) {
  e.preventDefault();
  const name = document.getElementById('station-borrow-name').value.trim();
  const due = document.getElementById('station-borrow-due').value;
  if (!name || !due) { alert('請填寫姓名與預計歸還日期'); return; }

  const btn = e.target.querySelector('button[type="submit"]');
  if (btn) { btn.disabled = true; btn.textContent = '🔄 處理中...'; }

  const closeAndRefresh = () => {
    const modal = document.getElementById('station-borrow-modal');
    if (modal) modal.style.display = 'none';
    loadTestStations();
  };

  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'borrowStation');
    url.searchParams.append('station', station);
    url.searchParams.append('borrower', name);
    url.searchParams.append('dt_due', due);
    const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    const text = await res.text();

    let result;
    try {
      result = JSON.parse(text);
    } catch (parseErr) {
      // GAS 轉址不穩回傳 HTML：伺服器通常已執行成功 → 關閉並刷新讓使用者看真實狀態
      console.warn('借用回應非 JSON（GAS 不穩），視為已送出:', text.slice(0, 120));
      closeAndRefresh();
      return;
    }

    if (result.success) {
      closeAndRefresh();
    } else if (result.error && result.error.indexOf(name) !== -1 && result.error.indexOf('借用中') !== -1) {
      // 已被「同一個人（你自己）」借用中 → 其實就是你剛剛借的，視為成功
      closeAndRefresh();
    } else {
      throw new Error(result.error || '借用失敗');
    }
  } catch (err) {
    alert('❌ 借用失敗：' + err.message);
    if (btn) { btn.disabled = false; btn.textContent = '✅ 確認借用'; }
  }
}

// 歸還測試站
async function returnStationBorrow(station) {
  if (!confirm(`確定要歸還測試站 ${station} 嗎？`)) return;
  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'returnStation');
    url.searchParams.append('station', station);
    const res = await fetch(url.toString(), { method: 'GET', redirect: 'follow' });
    const result = await res.json();
    if (result.success) {
      loadTestStations();
    } else {
      throw new Error(result.error || '歸還失敗');
    }
  } catch (err) {
    alert('❌ 歸還失敗：' + err.message);
  }
}

// 開啟測試站續借 Modal
function openStationPostponeModal(station, currentDue) {
  buildStationModal('station-postpone-modal', `
    <h2 style="color:#667eea;margin-bottom:20px;">⏰ 續借測試站 ${station}</h2>
    ${currentDue ? `<p style="color:#666;margin-bottom:12px;">目前預計歸還：<strong>${currentDue}</strong></p>` : ''}
    <form onsubmit="submitStationPostpone(event, '${station}')">
      <div class="form-group">
        <label>新的預計歸還日期 * <span style="font-size:0.85em;color:#888;">(以天為單位)</span></label>
        <input type="date" id="station-postpone-due" min="${stationMinDateTime().slice(0, 10)}" required style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ddd;border-radius:4px;">
      </div>
      <div style="text-align:center;margin-top:20px;">
        <button type="submit" class="btn-primary" style="background:#ffc107;color:#495057;">✅ 確認續借</button>
      </div>
    </form>
  `);
}

// 送出測試站續借
async function submitStationPostpone(e, station) {
  e.preventDefault();
  const newDue = document.getElementById('station-postpone-due').value;
  if (!newDue) { alert('請選擇新的預計歸還日期'); return; }

  const btn = e.target.querySelector('button[type="submit"]');
  if (btn) { btn.disabled = true; btn.textContent = '🔄 處理中...'; }

  try {
    const url = new URL(GAS_URL);
    url.searchParams.append('action', 'postponeStation');
    url.searchParams.append('station', station);
    url.searchParams.append('new_due_date', newDue);
    const result = await gasGetJson(url.toString());
    if (result.success) {
      const modal = document.getElementById('station-postpone-modal');
      if (modal) modal.style.display = 'none';
      loadTestStations();
    } else {
      throw new Error(result.error || '續借失敗');
    }
  } catch (err) {
    if (err.isGasGlitch) {
      const modal = document.getElementById('station-postpone-modal');
      if (modal) modal.style.display = 'none';
      loadTestStations();
      return;
    }
    alert('❌ 續借失敗：' + err.message);
    if (btn) { btn.disabled = false; btn.textContent = '✅ 確認續借'; }
  }
}
