// 家庭照顧與家用花費紀錄 APP 核心邏輯
(function () {
  "use strict";

  const STORAGE_KEY = "FAMILY_EXPENSE_DATA_V1";
  const GAS_URL_KEY = "FAMILY_EXPENSE_GAS_URL";
  const SPREADSHEET_ID = "1wut51zEYI7Ij0aBbCx-b7fpEA_n0lhdMht7PtEjnGHY";

  // Google Apps Script 代碼字串 (供一鍵複製)
  const GAS_CODE = `function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }
function handleRequest(e) {
  var output = { success: false, message: "" };
  try {
    var params = {};
    if (e.parameter && e.parameter.payload) {
      params = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      params = e.parameter;
    }
    var action = params.action || "ping";
    var category = params.category;
    var data = params.data || params;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (action === "ping") {
      output = { success: true, message: "Google Apps Script 雲端同步服務正常運作中！", time: new Date() };
    } else if (action === "add") {
      output = addRecord(ss, category, data);
    } else {
      output = { success: false, message: "未知的操作指令：" + action };
    }
  } catch (err) {
    output = { success: false, message: "執行錯誤：" + err.toString() };
  }
  var callback = (e.parameter && e.parameter.callback);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(output) + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.JSON);
}

function addRecord(ss, category, data) {
  var dateStr = formatDateToYYMMDD(data.date);
  var amount = Number(data.amount) || 0;
  var note = data.note || "";
  if (category === "medical") {
    var sheet = ss.getSheetByName("就醫、照顧花費記帳");
    if (!sheet) return { success: false, message: "找不到工作表「就醫、照顧花費記帳」" };
    var item = data.item || "其他";
    var type = "其他";
    if (/看診|住院/.test(item)) type = "醫療";
    else if (/長照|看護/.test(item)) type = "照護";
    var momAmt = "";
    if (note.indexOf("老媽") !== -1) momAmt = amount;
    var targetRow = findFirstEmptyRow(sheet, 2, 4);
    sheet.getRange(targetRow, 2).setValue(dateStr);
    sheet.getRange(targetRow, 3).setValue(item);
    sheet.getRange(targetRow, 4).setValue(type);
    sheet.getRange(targetRow, 5).setValue(amount);
    sheet.getRange(targetRow, 6).setValue(note);
    if (momAmt !== "") sheet.getRange(targetRow, 7).setValue(momAmt);
    return { success: true, message: "已寫入就醫照顧第 " + targetRow + " 列！", row: targetRow };
  } else if (category === "household") {
    var sheet = ss.getSheetByName("家用花費、繳款紀錄");
    if (!sheet) return { success: false, message: "找不到工作表「家用花費、繳款紀錄」" };
    var item = data.item || "其他";
    var account = "";
    var type = "其他";
    if (item === "房貸轉帳") { item = "轉帳"; account = "永豐"; type = "房貸"; }
    else if (/台電|北水|瓦斯/.test(item)) { type = "水電"; account = "竑郵局"; }
    else if (item === "中華電信") { type = "其他"; account = "竑郵局"; }
    var targetRow = findFirstEmptyRow(sheet, 2, 4);
    sheet.getRange(targetRow, 2).setValue(dateStr);
    sheet.getRange(targetRow, 3).setValue(item);
    sheet.getRange(targetRow, 4).setValue(account);
    sheet.getRange(targetRow, 5).setValue(type);
    sheet.getRange(targetRow, 6).setValue(amount);
    sheet.getRange(targetRow, 7).setValue(note);
    return { success: true, message: "已寫入家用繳款第 " + targetRow + " 列！", row: targetRow };
  } else if (category === "transfers") {
    var sheet = ss.getSheetByName("老哥轉帳記錄");
    if (!sheet) return { success: false, message: "找不到工作表「老哥轉帳記錄」" };
    var targetRow = findFirstEmptyRow(sheet, 2, 5);
    sheet.getRange(targetRow, 2).setValue(dateStr);
    sheet.getRange(targetRow, 7).setValue(amount);
    sheet.getRange(targetRow, 8).setValue(note);
    return { success: true, message: "已寫入老哥轉帳第 " + targetRow + " 列！", row: targetRow };
  }
  return { success: false, message: "不支援的類別" };
}

function formatDateToYYMMDD(val) {
  if (!val) {
    var now = new Date();
    var yy = String(now.getFullYear()).substring(2);
    var mm = String(now.getMonth() + 1).padStart(2, "0");
    var dd = String(now.getDate()).padStart(2, "0");
    return Number(yy + mm + dd);
  }
  var s = String(val).replace(/[-/]/g, "").trim();
  if (s.length === 8) return Number(s.substring(2));
  if (s.length === 6) return Number(s);
  return val;
}

function findFirstEmptyRow(sheet, colIndex, startRow) {
  var maxRows = sheet.getMaxRows();
  var values = sheet.getRange(startRow, colIndex, maxRows - startRow + 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v === "" || v === null || v === undefined) return startRow + i;
  }
  sheet.appendRow([""]);
  return sheet.getLastRow();
}`;

  // 資料結構狀態
  let appData = {
    medicalCare: [],
    household: [],
    transfers: []
  };

  // 當前篩選條件狀態
  const filters = {
    med: { keyword: "", item: "ALL", year: "ALL" },
    house: { keyword: "", item: "ALL", year: "ALL" },
    transfer: { keyword: "", year: "ALL" }
  };

  // 格式化千分位貨幣
  function formatMoney(num) {
    if (num === null || num === undefined || isNaN(num)) return "NT$ 0";
    return "NT$ " + Math.round(num).toLocaleString("en-US");
  }

  // 解析日期字串為 YYYY-MM-DD
  function parseDateString(v) {
    if (v === null || v === undefined) return null;
    const s = String(v).split('.')[0].trim();
    if (s.length === 6 && /^\d+$/.test(s)) {
      return `20${s.substring(0, 2)}-${s.substring(2, 4)}-${s.substring(4, 6)}`;
    } else if (s.length === 8 && /^\d+$/.test(s)) {
      return `${s.substring(0, 4)}-${s.substring(4, 6)}-${s.substring(6, 8)}`;
    } else if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      return s;
    }
    return null;
  }

  // 取得今日日期 (YYYY-MM-DD)
  function getTodayString() {
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, "0");
    const d = String(today.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  // Toast 提示
  function showToast(msg) {
    const toast = document.getElementById("toast");
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.add("show");
    setTimeout(() => {
      toast.classList.remove("show");
    }, 3200);
  }

  // 初始化資料載入
  function loadData() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        appData = JSON.parse(stored);
      } else if (window.INITIAL_DATA) {
        appData = JSON.parse(JSON.stringify(window.INITIAL_DATA));
        saveData();
      }
    } catch (e) {
      console.error("載入本機資料失敗，使用預設值:", e);
      if (window.INITIAL_DATA) {
        appData = JSON.parse(JSON.stringify(window.INITIAL_DATA));
      }
    }
    updateSyncBadge();
  }

  // 儲存資料至 LocalStorage
  function saveData() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(appData));
    } catch (e) {
      console.error("無法寫入 LocalStorage:", e);
      showToast("儲存失敗：本機空間不足");
    }
  }

  function updateSyncBadge() {
    const badge = document.getElementById("syncBadge");
    const gasUrl = localStorage.getItem(GAS_URL_KEY);
    if (badge) {
      if (gasUrl && gasUrl.trim().startsWith("http")) {
        badge.textContent = `⚡ 雲端寫入同步中`;
        badge.style.background = "#dcfce7";
        badge.style.color = "#15803d";
      } else {
        badge.textContent = `僅本機儲存 (未設同步網址)`;
        badge.style.background = "#f1f5f9";
        badge.style.color = "#64748b";
      }
    }
  }

  // 雲端寫入通訊器 (支援 JSONP 與 fetch fallback)
  function sendToGAS(gasUrl, payload) {
    return new Promise((resolve, reject) => {
      const cbName = "gasSyncCb_" + Math.random().toString(36).substring(2, 9);
      const script = document.createElement("script");
      let finished = false;

      const timer = setTimeout(() => {
        if (!finished) {
          finished = true;
          cleanup();
          // 若 JSONP 逾時，嘗試以 no-cors POST 發送
          fallbackPost(gasUrl, payload).then(resolve).catch(reject);
        }
      }, 7000);

      function cleanup() {
        clearTimeout(timer);
        delete window[cbName];
        if (script.parentNode) script.parentNode.removeChild(script);
      }

      window[cbName] = function (res) {
        if (!finished) {
          finished = true;
          cleanup();
          resolve(res);
        }
      };

      const query = `payload=${encodeURIComponent(JSON.stringify(payload))}&callback=${cbName}&_t=${Date.now()}`;
      script.src = gasUrl + (gasUrl.includes("?") ? "&" : "?") + query;
      script.onerror = function () {
        if (!finished) {
          finished = true;
          cleanup();
          fallbackPost(gasUrl, payload).then(resolve).catch(reject);
        }
      };
      document.head.appendChild(script);
    });
  }

  function fallbackPost(gasUrl, payload) {
    return fetch(gasUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(payload)
    }).then(() => {
      return { success: true, message: "已送出更新至雲端試算表！" };
    });
  }

  // 自動同步單筆寫入至 Google 雲端試算表
  async function syncRecordToCloud(category, record) {
    const gasUrl = (localStorage.getItem(GAS_URL_KEY) || "").trim();
    if (!gasUrl) {
      showToast("💾 已儲存！(若需手機自動寫入 Google 試算表，請在「⚙️ 雲端設定」填入同步網址)");
      return;
    }

    showToast("☁️ 正在即時同步寫入 Google 雲端試算表...");

    try {
      const res = await sendToGAS(gasUrl, {
        action: "add",
        category: category,
        data: record
      });

      if (res && res.success !== false) {
        showToast("✅ 已成功寫入 Google 雲端試算表！其他人已可查看最新數字！");
      } else {
        showToast("⚠️ 本機已儲存，雲端回應：" + (res.message || "請檢查設定"));
      }
    } catch (err) {
      console.warn("雲端同步連線異常:", err);
      showToast("⚠️ 本機已儲存，雲端同步失敗（請確認連線或設定網址）");
    }
  }

  // 標籤頁面切換
  function setupNavigation() {
    const desktopBtns = document.querySelectorAll(".desktop-nav .nav-tab-btn");
    const bottomBtns = document.querySelectorAll(".bottom-nav .bottom-tab-btn");

    function switchTab(tabId) {
      desktopBtns.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.tab === tabId);
      });
      bottomBtns.forEach(btn => {
        btn.classList.toggle("active", btn.dataset.tab === tabId);
      });

      document.querySelectorAll(".tab-panel").forEach(panel => {
        panel.classList.toggle("active", panel.id === `tab-${tabId}`);
      });

      if (tabId === "dashboard") renderDashboard();
      if (tabId === "medical") renderMedicalList();
      if (tabId === "household") renderHouseholdList();
      if (tabId === "transfers") renderTransferList();

      window.scrollTo({ top: 0, behavior: "smooth" });
    }

    desktopBtns.forEach(btn => {
      btn.addEventListener("click", () => switchTab(btn.dataset.tab));
    });
    bottomBtns.forEach(btn => {
      btn.addEventListener("click", () => switchTab(btn.dataset.tab));
    });
  }

  // 表單摺疊切換
  function setupFormToggles() {
    const btnMed = document.getElementById("btnToggleMedForm");
    const cardMed = document.getElementById("formMedCard");
    const btnCancelMed = document.getElementById("btnCancelMed");
    btnMed.addEventListener("click", () => cardMed.classList.toggle("open"));
    btnCancelMed.addEventListener("click", () => cardMed.classList.remove("open"));

    const btnHouse = document.getElementById("btnToggleHouseForm");
    const cardHouse = document.getElementById("formHouseCard");
    const btnCancelHouse = document.getElementById("btnCancelHouse");
    btnHouse.addEventListener("click", () => cardHouse.classList.toggle("open"));
    btnCancelHouse.addEventListener("click", () => cardHouse.classList.remove("open"));

    const btnTransfer = document.getElementById("btnToggleTransferForm");
    const cardTransfer = document.getElementById("formTransferCard");
    const btnCancelTransfer = document.getElementById("btnCancelTransfer");
    btnTransfer.addEventListener("click", () => cardTransfer.classList.toggle("open"));
    btnCancelTransfer.addEventListener("click", () => cardTransfer.classList.remove("open"));

    const today = getTodayString();
    document.getElementById("medDate").value = today;
    document.getElementById("houseDate").value = today;
    document.getElementById("transferDate").value = today;
  }

  // 1. 渲染儀表板 (Dashboard)
  function renderDashboard() {
    const totalTransfers = appData.transfers.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const totalMed = appData.medicalCare.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const totalHouse = appData.household.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const totalExpenses = totalMed + totalHouse;
    const balance = totalTransfers - totalExpenses;

    document.getElementById("kpiTransfer").textContent = formatMoney(totalTransfers);
    document.getElementById("kpiTransferCount").textContent = `共 ${appData.transfers.length} 筆轉入`;

    document.getElementById("kpiMedical").textContent = formatMoney(totalMed);
    document.getElementById("kpiMedicalCount").textContent = `共 ${appData.medicalCare.length} 筆支出 (含老媽)`;

    document.getElementById("kpiHousehold").textContent = formatMoney(totalHouse);
    document.getElementById("kpiHouseholdCount").textContent = `共 ${appData.household.length} 筆日常繳款`;

    const elBalance = document.getElementById("kpiBalance");
    elBalance.textContent = formatMoney(balance);
    elBalance.style.color = balance >= 0 ? "var(--primary-dark)" : "var(--danger)";

    renderYearlyChart();
    renderCategoryRanking();
    renderRecentActivities();
  }

  // 儀表板：年度收支趨勢長條圖
  function renderYearlyChart() {
    const chartContainer = document.getElementById("yearlyChart");
    if (!chartContainer) return;
    chartContainer.innerHTML = "";

    const years = ["2022", "2023", "2024", "2025", "2026"];
    const stats = {};
    years.forEach(y => {
      stats[y] = { inflow: 0, care: 0, house: 0 };
    });

    appData.transfers.forEach(r => {
      const y = (r.date || "").substring(0, 4);
      if (stats[y]) stats[y].inflow += Number(r.amount) || 0;
    });
    appData.medicalCare.forEach(r => {
      const y = (r.date || "").substring(0, 4);
      if (stats[y]) stats[y].care += Number(r.amount) || 0;
    });
    appData.household.forEach(r => {
      const y = (r.date || "").substring(0, 4);
      if (stats[y]) stats[y].house += Number(r.amount) || 0;
    });

    let maxVal = 100000;
    years.forEach(y => {
      maxVal = Math.max(maxVal, stats[y].inflow, stats[y].care + stats[y].house);
    });

    years.forEach(y => {
      const totalExpense = stats[y].care + stats[y].house;
      const inflow = stats[y].inflow;
      const careWidth = Math.min(100, Math.round((stats[y].care / maxVal) * 100));
      const houseWidth = Math.min(100, Math.round((stats[y].house / maxVal) * 100));
      const inflowWidth = Math.min(100, Math.round((inflow / maxVal) * 100));

      const row = document.createElement("div");
      row.className = "bar-row";
      row.innerHTML = `
        <div class="bar-label">${y}年</div>
        <div style="flex: 1; display: flex; flex-direction: column; gap: 4px;">
          <!-- 轉入條 -->
          <div style="display: flex; align-items: center; gap: 6px;">
            <span style="font-size: 0.72rem; color: var(--success); width: 32px;">轉入</span>
            <div class="bar-track" style="height: 14px;">
              <div class="bar-fill bar-inflow" style="width: ${inflowWidth}%;"></div>
            </div>
            <span style="font-size: 0.75rem; font-weight: 700; width: 80px; text-align: right; color: var(--success);">${formatMoney(inflow)}</span>
          </div>
          <!-- 支出條 -->
          <div style="display: flex; align-items: center; gap: 6px;">
            <span style="font-size: 0.72rem; color: var(--slate-600); width: 32px;">支出</span>
            <div class="bar-track" style="height: 14px;">
              <div class="bar-fill bar-care" style="width: ${careWidth}%;" title="就醫照護"></div>
              <div class="bar-fill bar-house" style="width: ${houseWidth}%;" title="家用生活"></div>
            </div>
            <span style="font-size: 0.75rem; font-weight: 700; width: 80px; text-align: right; color: var(--slate-700);">${formatMoney(totalExpense)}</span>
          </div>
        </div>
      `;
      chartContainer.appendChild(row);
    });
  }

  // 儀表板：類別排行
  function renderCategoryRanking() {
    const container = document.getElementById("categoryRanking");
    if (!container) return;
    container.innerHTML = "";

    const catMap = {};
    appData.medicalCare.forEach(r => {
      const item = r.item || "其他";
      catMap[item] = (catMap[item] || 0) + (Number(r.amount) || 0);
    });
    appData.household.forEach(r => {
      const item = r.item || "其他";
      catMap[item] = (catMap[item] || 0) + (Number(r.amount) || 0);
    });

    const sorted = Object.entries(catMap).sort((a, b) => b[1] - a[1]).slice(0, 6);
    const totalExp = appData.medicalCare.reduce((s, x) => s + (Number(x.amount)||0), 0) +
                     appData.household.reduce((s, x) => s + (Number(x.amount)||0), 0);

    sorted.forEach(([name, amt]) => {
      const pct = totalExp > 0 ? ((amt / totalExp) * 100).toFixed(1) : 0;
      const itemDiv = document.createElement("div");
      itemDiv.className = "breakdown-item";
      itemDiv.innerHTML = `
        <div style="display: flex; align-items: center; gap: 8px;">
          <span class="badge badge-${name}">${name}</span>
          <span style="font-size: 0.8rem; color: var(--slate-500);">${pct}%</span>
        </div>
        <div style="font-weight: 700; color: var(--slate-800);">${formatMoney(amt)}</div>
      `;
      container.appendChild(itemDiv);
    });
  }

  // 儀表板：最近動態
  function renderRecentActivities() {
    const container = document.getElementById("recentActivityList");
    if (!container) return;
    container.innerHTML = "";

    const combined = [];
    appData.medicalCare.forEach(r => combined.push({ ...r, type: "medical", typeName: "就醫照顧" }));
    appData.household.forEach(r => combined.push({ ...r, type: "household", typeName: "家用繳款" }));
    appData.transfers.forEach(r => combined.push({ ...r, type: "transfer", item: "轉帳", typeName: "老哥轉帳" }));

    combined.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
    const recent = combined.slice(0, 5);

    recent.forEach(r => {
      const isTransfer = r.type === "transfer";
      const card = document.createElement("div");
      card.className = "record-card";
      card.innerHTML = `
        <div class="record-top">
          <div class="record-date">📅 ${r.date}</div>
          <span class="badge badge-${r.item || '轉帳'}">${r.item || '轉帳'}</span>
        </div>
        <div class="record-mid">
          <div class="record-item-name" style="font-size: 0.92rem; color: var(--slate-600);">${r.typeName}</div>
          <div class="record-amount ${isTransfer ? 'inflow' : ''}">${isTransfer ? '+' : '-'}${formatMoney(r.amount)}</div>
        </div>
        ${r.note ? `<div class="record-note">${r.note}</div>` : ""}
      `;
      container.appendChild(card);
    });
  }

  // 2. 渲染就醫、照顧花費清單
  function renderMedicalList() {
    const container = document.getElementById("medRecordList");
    if (!container) return;
    container.innerHTML = "";

    const keyword = filters.med.keyword.toLowerCase().trim();
    const itemFilter = filters.med.item;
    const yearFilter = filters.med.year;

    const filtered = appData.medicalCare.filter(r => {
      if (itemFilter !== "ALL" && r.item !== itemFilter) return false;
      if (yearFilter !== "ALL" && !(r.date || "").startsWith(yearFilter)) return false;
      if (keyword) {
        const text = `${r.date} ${r.item} ${r.note || ""}`.toLowerCase();
        if (!text.includes(keyword)) return false;
      }
      return true;
    });

    filtered.sort((a, b) => (b.date || "").localeCompare(a.date || ""));

    const subtotal = filtered.reduce((s, x) => s + (Number(x.amount) || 0), 0);
    document.getElementById("medResultCount").textContent = filtered.length;
    document.getElementById("medResultTotal").textContent = formatMoney(subtotal);

    if (filtered.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">📋</div>
          <div>未找到相符的就醫照顧紀錄</div>
        </div>
      `;
      return;
    }

    filtered.forEach(r => {
      const card = document.createElement("div");
      card.className = "record-card";
      
      const hasMom = (r.note || "").includes("老媽");
      let displayNote = r.note || "";
      if (hasMom) {
        displayNote = `<span class="mom-tag">老媽款項</span> ` + displayNote.replace(/老媽/g, "").replace(/^;\s*/, "").replace(/;\s*$/, "").trim();
      }

      card.innerHTML = `
        <div class="record-top">
          <div class="record-date">📅 ${r.date}</div>
          <span class="badge badge-${r.item}">${r.item}</span>
        </div>
        <div class="record-mid">
          <div class="record-item-name">${r.item}</div>
          <div class="record-amount">${formatMoney(r.amount)}</div>
        </div>
        ${r.note ? `<div class="record-note">${displayNote}</div>` : ""}
        <div class="record-footer">
          <button class="btn btn-secondary btn-sm btn-edit" data-category="medical" data-id="${r.id}">✏️ 編輯</button>
          <button class="btn btn-danger btn-sm btn-delete" data-category="medical" data-id="${r.id}">🗑️ 刪除</button>
        </div>
      `;
      container.appendChild(card);
    });

    attachItemActions(container);
  }

  // 3. 渲染家用繳款紀錄清單
  function renderHouseholdList() {
    const container = document.getElementById("houseRecordList");
    if (!container) return;
    container.innerHTML = "";

    const keyword = filters.house.keyword.toLowerCase().trim();
    const itemFilter = filters.house.item;
    const yearFilter = filters.house.year;

    const filtered = appData.household.filter(r => {
      if (itemFilter !== "ALL" && r.item !== itemFilter) return false;
      if (yearFilter !== "ALL" && !(r.date || "").startsWith(yearFilter)) return false;
      if (keyword) {
        const text = `${r.date} ${r.item} ${r.note || ""}`.toLowerCase();
        if (!text.includes(keyword)) return false;
      }
      return true;
    });

    filtered.sort((a, b) => (b.date || "").localeCompare(a.date || ""));

    const subtotal = filtered.reduce((s, x) => s + (Number(x.amount) || 0), 0);
    document.getElementById("houseResultCount").textContent = filtered.length;
    document.getElementById("houseResultTotal").textContent = formatMoney(subtotal);

    if (filtered.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">🏠</div>
          <div>未找到相符的家用繳款紀錄</div>
        </div>
      `;
      return;
    }

    filtered.forEach(r => {
      const card = document.createElement("div");
      card.className = "record-card";
      card.innerHTML = `
        <div class="record-top">
          <div class="record-date">📅 ${r.date}</div>
          <span class="badge badge-${r.item}">${r.item}</span>
        </div>
        <div class="record-mid">
          <div class="record-item-name">${r.item}</div>
          <div class="record-amount">${formatMoney(r.amount)}</div>
        </div>
        ${r.note ? `<div class="record-note">${r.note}</div>` : ""}
        <div class="record-footer">
          <button class="btn btn-secondary btn-sm btn-edit" data-category="household" data-id="${r.id}">✏️ 編輯</button>
          <button class="btn btn-danger btn-sm btn-delete" data-category="household" data-id="${r.id}">🗑️ 刪除</button>
        </div>
      `;
      container.appendChild(card);
    });

    attachItemActions(container);
  }

  // 4. 渲染老哥轉帳紀錄清單
  function renderTransferList() {
    const container = document.getElementById("transferRecordList");
    if (!container) return;
    container.innerHTML = "";

    const keyword = filters.transfer.keyword.toLowerCase().trim();
    const yearFilter = filters.transfer.year;

    const filtered = appData.transfers.filter(r => {
      if (yearFilter !== "ALL" && !(r.date || "").startsWith(yearFilter)) return false;
      if (keyword) {
        const text = `${r.date} ${r.amount} ${r.note || ""}`.toLowerCase();
        if (!text.includes(keyword)) return false;
      }
      return true;
    });

    filtered.sort((a, b) => (b.date || "").localeCompare(a.date || ""));

    const subtotal = filtered.reduce((s, x) => s + (Number(x.amount) || 0), 0);
    document.getElementById("transferResultCount").textContent = filtered.length;
    document.getElementById("transferResultTotal").textContent = formatMoney(subtotal);

    if (filtered.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">💰</div>
          <div>未找到相符的轉帳紀錄</div>
        </div>
      `;
      return;
    }

    filtered.forEach(r => {
      const card = document.createElement("div");
      card.className = "record-card";
      card.innerHTML = `
        <div class="record-top">
          <div class="record-date">📅 日期：${r.date}</div>
          <span class="badge badge-轉帳">入帳匯款</span>
        </div>
        <div class="record-mid">
          <div class="record-item-name">換匯後金額 NTD</div>
          <div class="record-amount inflow">+${formatMoney(r.amount)}</div>
        </div>
        ${r.note ? `<div class="record-note">備註：${r.note}</div>` : ""}
        <div class="record-footer">
          <button class="btn btn-secondary btn-sm btn-edit" data-category="transfers" data-id="${r.id}">✏️ 編輯</button>
          <button class="btn btn-danger btn-sm btn-delete" data-category="transfers" data-id="${r.id}">🗑️ 刪除</button>
        </div>
      `;
      container.appendChild(card);
    });

    attachItemActions(container);
  }

  // 綁定卡片上的「編輯」與「刪除」按鈕
  function attachItemActions(container) {
    container.querySelectorAll(".btn-edit").forEach(btn => {
      btn.addEventListener("click", () => {
        openEditModal(btn.dataset.category, btn.dataset.id);
      });
    });

    container.querySelectorAll(".btn-delete").forEach(btn => {
      btn.addEventListener("click", () => {
        deleteRecord(btn.dataset.category, btn.dataset.id);
      });
    });
  }

  // 刪除單筆紀錄
  function deleteRecord(category, id) {
    if (!confirm("確定要刪除這筆紀錄嗎？")) return;
    const list = appData[category];
    const index = list.findIndex(x => x.id === id);
    if (index !== -1) {
      list.splice(index, 1);
      saveData();
      showToast("已從本機刪除紀錄");
      if (category === "medical") renderMedicalList();
      else if (category === "household") renderHouseholdList();
      else if (category === "transfers") renderTransferList();
      renderDashboard();
    }
  }

  // 開啟編輯 Modal
  function openEditModal(category, id) {
    const list = appData[category];
    const record = list.find(x => x.id === id);
    if (!record) return;

    document.getElementById("editCategory").value = category;
    document.getElementById("editId").value = id;
    document.getElementById("editDate").value = record.date || getTodayString();
    document.getElementById("editAmount").value = record.amount || 0;
    document.getElementById("editNote").value = record.note || "";

    const editItemGroup = document.getElementById("editItemGroup");
    const editItem = document.getElementById("editItem");
    editItem.innerHTML = "";

    if (category === "medical") {
      editItemGroup.style.display = "block";
      const medOptions = ["看診", "長照費", "住院費", "看護費", "其他"];
      medOptions.forEach(opt => {
        const o = document.createElement("option");
        o.value = opt;
        o.textContent = opt;
        if (record.item === opt) o.selected = true;
        editItem.appendChild(o);
      });
      document.getElementById("editModalTitle").textContent = "✏️ 編輯就醫照顧紀錄";
    } else if (category === "household") {
      editItemGroup.style.display = "block";
      const houseOptions = ["台電", "中華電信", "瓦斯", "北水", "房貸轉帳", "墓園管理費", "其他"];
      houseOptions.forEach(opt => {
        const o = document.createElement("option");
        o.value = opt;
        o.textContent = opt;
        if (record.item === opt) o.selected = true;
        editItem.appendChild(o);
      });
      document.getElementById("editModalTitle").textContent = "✏️ 編輯家用繳款紀錄";
    } else if (category === "transfers") {
      editItemGroup.style.display = "none";
      document.getElementById("editModalTitle").textContent = "✏️ 編輯老哥轉帳紀錄";
    }

    document.getElementById("editModal").classList.add("open");
  }

  function closeEditModal() {
    document.getElementById("editModal").classList.remove("open");
  }

  // 綁定篩選器監聽
  function setupFilters() {
    const searchMed = document.getElementById("searchMed");
    const filterMedItem = document.getElementById("filterMedItem");
    const filterMedYear = document.getElementById("filterMedYear");
    searchMed.addEventListener("input", (e) => {
      filters.med.keyword = e.target.value;
      renderMedicalList();
    });
    filterMedItem.addEventListener("change", (e) => {
      filters.med.item = e.target.value;
      renderMedicalList();
    });
    filterMedYear.addEventListener("change", (e) => {
      filters.med.year = e.target.value;
      renderMedicalList();
    });

    const searchHouse = document.getElementById("searchHouse");
    const filterHouseItem = document.getElementById("filterHouseItem");
    const filterHouseYear = document.getElementById("filterHouseYear");
    searchHouse.addEventListener("input", (e) => {
      filters.house.keyword = e.target.value;
      renderHouseholdList();
    });
    filterHouseItem.addEventListener("change", (e) => {
      filters.house.item = e.target.value;
      renderHouseholdList();
    });
    filterHouseYear.addEventListener("change", (e) => {
      filters.house.year = e.target.value;
      renderHouseholdList();
    });

    const searchTransfer = document.getElementById("searchTransfer");
    const filterTransferYear = document.getElementById("filterTransferYear");
    searchTransfer.addEventListener("input", (e) => {
      filters.transfer.keyword = e.target.value;
      renderTransferList();
    });
    filterTransferYear.addEventListener("change", (e) => {
      filters.transfer.year = e.target.value;
      renderTransferList();
    });
  }

  // 綁定表單送出（新增與編輯）
  function setupForms() {
    // 1. 新增就醫紀錄 (手機更新 -> 本機儲存 + 自動寫入 Google 試算表)
    document.getElementById("formMed").addEventListener("submit", (e) => {
      e.preventDefault();
      const date = document.getElementById("medDate").value;
      const item = document.getElementById("medItem").value;
      const amount = Number(document.getElementById("medAmount").value);
      const note = document.getElementById("medNote").value.trim();

      const newRecord = {
        id: `med_${Date.now()}`,
        date,
        item,
        amount,
        note
      };

      appData.medicalCare.unshift(newRecord);
      saveData();
      document.getElementById("medAmount").value = "";
      document.getElementById("medNote").value = "";
      document.getElementById("formMedCard").classList.remove("open");
      renderMedicalList();
      renderDashboard();

      // 同步寫入 Google 試算表
      syncRecordToCloud("medical", newRecord);
    });

    // 2. 新增家用紀錄 (手機更新 -> 本機儲存 + 自動寫入 Google 試算表)
    document.getElementById("formHouse").addEventListener("submit", (e) => {
      e.preventDefault();
      const date = document.getElementById("houseDate").value;
      const item = document.getElementById("houseItem").value;
      const amount = Number(document.getElementById("houseAmount").value);
      const note = document.getElementById("houseNote").value.trim();

      const newRecord = {
        id: `house_${Date.now()}`,
        date,
        item,
        amount,
        note
      };

      appData.household.unshift(newRecord);
      saveData();
      document.getElementById("houseAmount").value = "";
      document.getElementById("houseNote").value = "";
      document.getElementById("formHouseCard").classList.remove("open");
      renderHouseholdList();
      renderDashboard();

      syncRecordToCloud("household", newRecord);
    });

    // 3. 新增老哥轉帳紀錄 (手機更新 -> 本機儲存 + 自動寫入 Google 試算表)
    document.getElementById("formTransfer").addEventListener("submit", (e) => {
      e.preventDefault();
      const date = document.getElementById("transferDate").value;
      const amount = Number(document.getElementById("transferAmount").value);
      const note = document.getElementById("transferNote").value.trim();

      const newRecord = {
        id: `transfer_${Date.now()}`,
        date,
        amount,
        note
      };

      appData.transfers.unshift(newRecord);
      saveData();
      document.getElementById("transferAmount").value = "";
      document.getElementById("transferNote").value = "";
      document.getElementById("formTransferCard").classList.remove("open");
      renderTransferList();
      renderDashboard();

      syncRecordToCloud("transfers", newRecord);
    });

    // 4. 編輯表單送出
    document.getElementById("formEdit").addEventListener("submit", (e) => {
      e.preventDefault();
      const category = document.getElementById("editCategory").value;
      const id = document.getElementById("editId").value;
      const date = document.getElementById("editDate").value;
      const amount = Number(document.getElementById("editAmount").value);
      const note = document.getElementById("editNote").value.trim();

      const list = appData[category];
      const record = list.find(x => x.id === id);
      if (record) {
        record.date = date;
        record.amount = amount;
        record.note = note;
        if (category !== "transfers") {
          record.item = document.getElementById("editItem").value;
        }
        saveData();
        showToast("已更新紀錄！");
        closeEditModal();
        if (category === "medical") renderMedicalList();
        else if (category === "household") renderHouseholdList();
        else if (category === "transfers") renderTransferList();
        renderDashboard();
      }
    });

    document.getElementById("btnCancelEdit").addEventListener("click", closeEditModal);
    document.getElementById("btnCloseEditModal").addEventListener("click", closeEditModal);
  }

  // Google 雲端試算表拉取器
  function fetchSheetDataJSONP(sheetName) {
    return new Promise((resolve, reject) => {
      const cbName = "gvizCb_" + Math.random().toString(36).substring(2, 9);
      const script = document.createElement("script");
      const timer = setTimeout(() => {
        cleanup();
        reject(new Error("連線逾時"));
      }, 15000);

      function cleanup() {
        clearTimeout(timer);
        delete window[cbName];
        if (script.parentNode) script.parentNode.removeChild(script);
      }

      window[cbName] = function (res) {
        cleanup();
        if (res && res.table && res.table.rows) {
          resolve(res.table.rows);
        } else {
          reject(new Error("試算表回傳格式錯誤"));
        }
      };

      script.src = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=responseHandler:${cbName}&sheet=${encodeURIComponent(sheetName)}&t=${Date.now()}`;
      script.onerror = function () {
        cleanup();
        reject(new Error(`無法連線至工作表「${sheetName}」`));
      };
      document.head.appendChild(script);
    });
  }

  // 從 Google 試算表拉取最新數據
  async function pullFromGoogleSheets() {
    const btnSync = document.getElementById("btnSyncCloud");
    const btnModalSync = document.getElementById("btnModalSyncCloud");
    const syncIcon = btnSync ? btnSync.querySelector(".sync-icon") : null;

    if (syncIcon) syncIcon.classList.add("spinning");
    if (btnSync) btnSync.disabled = true;
    if (btnModalSync) btnModalSync.disabled = true;

    showToast("⏳ 正在從 Google 雲端試算表讀取最新資料...");

    try {
      const [rowsMed, rowsHouse, rowsTransfer] = await Promise.all([
        fetchSheetDataJSONP("就醫、照顧花費記帳"),
        fetchSheetDataJSONP("家用花費、繳款紀錄"),
        fetchSheetDataJSONP("老哥轉帳記錄")
      ]);

      // 1. 就醫照顧
      const newMed = [];
      let medId = 1;
      for (const r of rowsMed) {
        const c = r.c || [];
        if (c.length < 5 || !c[1]) continue;
        const d = parseDateString(c[1].v);
        if (!d) continue;

        const rawItem = String(c[2] && c[2].v ? c[2].v : "").trim();
        let amt = Number(c[4] && c[4].v !== undefined ? c[4].v : 0) || 0;
        const note = String(c[5] && c[5].v ? c[5].v : "").trim();
        const mom = c[6] && c[6].v !== undefined ? c[6].v : null;

        const notes = note ? [note] : [];
        if (mom !== null && mom !== undefined && String(mom).trim() !== "") {
          const momNum = Number(mom);
          if (!isNaN(momNum)) {
            amt += momNum;
            notes.push("老媽");
          } else {
            notes.push(`老媽: ${mom}`);
          }
        }

        let itemCat = "其他";
        if (/看診|門診|回診|急診/.test(rawItem)) {
          itemCat = "看診";
          if (rawItem !== "看診") notes.unshift(`原項目: ${rawItem}`);
        } else if (/長照/.test(rawItem)) {
          itemCat = "長照費";
          if (rawItem !== "長照費") notes.unshift(`原項目: ${rawItem}`);
        } else if (/住院|急診預繳/.test(rawItem)) {
          itemCat = "住院費";
          if (rawItem !== "住院費") notes.unshift(`原項目: ${rawItem}`);
        } else if (/看護/.test(rawItem)) {
          itemCat = "看護費";
          if (rawItem !== "看護費") notes.unshift(`原項目: ${rawItem}`);
        } else {
          itemCat = "其他";
          if (rawItem && rawItem !== "其他") notes.unshift(`原項目: ${rawItem}`);
        }

        newMed.push({
          id: `med_${medId++}`,
          date: d,
          item: itemCat,
          amount: Math.round(amt),
          note: notes.join("; ")
        });
      }

      // 2. 家用花費
      const newHouse = [];
      let houseId = 1;
      for (const r of rowsHouse) {
        const c = r.c || [];
        if (c.length < 6 || !c[1]) continue;
        const d = parseDateString(c[1].v);
        if (!d) continue;

        const rawItem = String(c[2] && c[2].v ? c[2].v : "").trim();
        const amt = Number(c[5] && c[5].v !== undefined ? c[5].v : 0) || 0;
        const note = String(c[6] && c[6].v ? c[6].v : "").trim();
        const notes = note ? [note] : [];

        let itemCat = "其他";
        if (rawItem.includes("台電")) itemCat = "台電";
        else if (rawItem.includes("中華電信")) itemCat = "中華電信";
        else if (rawItem.includes("瓦斯")) itemCat = "瓦斯";
        else if (rawItem.includes("北水") || rawItem.includes("水")) itemCat = "北水";
        else if (rawItem.includes("轉帳")) itemCat = "房貸轉帳";
        else if (rawItem.includes("墓園管理費")) {
          itemCat = "墓園管理費";
          if (rawItem.includes("土城")) notes.unshift("土城");
          else if (rawItem.includes("八里")) notes.unshift("八里");
        } else {
          itemCat = "其他";
          if (rawItem && rawItem !== "其他") notes.unshift(`原項目: ${rawItem}`);
        }

        newHouse.push({
          id: `house_${houseId++}`,
          date: d,
          item: itemCat,
          amount: Math.round(amt),
          note: notes.join("; ")
        });
      }

      // 3. 老哥轉帳
      const newTransfer = [];
      let transferId = 1;
      for (const r of rowsTransfer) {
        const c = r.c || [];
        if (c.length < 7 || !c[1]) continue;
        const d = parseDateString(c[1].v);
        if (!d) continue;

        const twd = c[6] && c[6].v !== undefined ? c[6].v : null;
        const note = String(c[7] && c[7].v ? c[7].v : "").trim();
        if (twd !== null && String(twd).trim() !== "") {
          const amt = Number(twd);
          if (!isNaN(amt)) {
            newTransfer.push({
              id: `transfer_${transferId++}`,
              date: d,
              amount: Math.round(amt),
              note
            });
          }
        }
      }

      appData = {
        medicalCare: newMed,
        household: newHouse,
        transfers: newTransfer
      };

      saveData();
      renderDashboard();
      renderMedicalList();
      renderHouseholdList();
      renderTransferList();

      showToast("🎉 已成功拉取 Google 雲端試算表最新資料！");
    } catch (err) {
      console.error("雲端讀取失敗:", err);
      showToast("❌ 雲端讀取失敗：" + err.message);
    } finally {
      if (syncIcon) syncIcon.classList.remove("spinning");
      if (btnSync) btnSync.disabled = false;
      if (btnModalSync) btnModalSync.disabled = false;
    }
  }

  // 設定與備份管理
  function setupBackupAndExport() {
    const backupModal = document.getElementById("backupModal");
    document.getElementById("btnBackup").addEventListener("click", () => {
      const inputGas = document.getElementById("inputGasUrl");
      if (inputGas) inputGas.value = localStorage.getItem(GAS_URL_KEY) || "";
      updateSyncBadge();
      backupModal.classList.add("open");
    });
    document.getElementById("btnCloseBackupModal").addEventListener("click", () => {
      backupModal.classList.remove("open");
    });

    // 儲存 Google Apps Script 同步網址
    document.getElementById("btnSaveGasUrl").addEventListener("click", () => {
      const url = (document.getElementById("inputGasUrl").value || "").trim();
      localStorage.setItem(GAS_URL_KEY, url);
      updateSyncBadge();
      showToast(url ? "✅ 雲端同步網址已儲存！手機記帳將自動寫入試算表。" : "已清除同步網址（僅本機儲存）");
    });

    // 測試連線
    document.getElementById("btnTestGasUrl").addEventListener("click", async () => {
      const url = (document.getElementById("inputGasUrl").value || "").trim();
      if (!url) {
        alert("請先輸入 Google Apps Script 網頁應用程式網址！");
        return;
      }
      showToast("⚡ 正在測試與 Google 雲端試算表連線...");
      try {
        const res = await sendToGAS(url, { action: "ping" });
        alert("🎉 連線成功！\n" + (res.message || "Google Apps Script 雲端同步已就緒！"));
      } catch (e) {
        alert("連線失敗：" + e.message + "\n請確認網址正確且部署設定為「所有人 (Anyone) 具存取權」");
      }
    });

    // 複製 Google Apps Script 程式碼
    document.getElementById("btnCopyGasCode").addEventListener("click", () => {
      navigator.clipboard.writeText(GAS_CODE).then(() => {
        showToast("📋 已複製後端腳本！請到 Google 試算表 Apps Script 貼上。");
      }).catch(() => {
        // Fallback
        const ta = document.createElement("textarea");
        ta.value = GAS_CODE;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        showToast("📋 已複製後端腳本！");
      });
    });

    // 雲端拉取按鈕綁定
    const btnSyncCloud = document.getElementById("btnSyncCloud");
    if (btnSyncCloud) btnSyncCloud.addEventListener("click", pullFromGoogleSheets);

    const btnModalSyncCloud = document.getElementById("btnModalSyncCloud");
    if (btnModalSyncCloud) btnModalSyncCloud.addEventListener("click", pullFromGoogleSheets);

    // 匯出 CSV (Excel UTF-8 with BOM)
    function exportToCSV() {
      let csvContent = "\uFEFF";
      csvContent += "類別,日期,項目,金額NTD,備註\r\n";

      appData.medicalCare.forEach(r => {
        csvContent += `"就醫照護","${r.date}","${r.item}",${r.amount},"${(r.note||"").replace(/"/g, '""')}"\r\n`;
      });
      appData.household.forEach(r => {
        csvContent += `"家用繳款","${r.date}","${r.item}",${r.amount},"${(r.note||"").replace(/"/g, '""')}"\r\n`;
      });
      appData.transfers.forEach(r => {
        csvContent += `"老哥轉帳","${r.date}","換匯後金額",${r.amount},"${(r.note||"").replace(/"/g, '""')}"\r\n`;
      });

      const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `家庭照護與家用開銷紀錄_${getTodayString()}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast("已成功匯出 CSV 試算表檔案！");
    }

    document.getElementById("btnExport").addEventListener("click", exportToCSV);
    document.getElementById("btnExportCSV").addEventListener("click", exportToCSV);

    // 下載 JSON 備份檔
    document.getElementById("btnDownloadBackup").addEventListener("click", () => {
      const jsonStr = JSON.stringify(appData, null, 2);
      const blob = new Blob([jsonStr], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `家庭開銷備份_${getTodayString()}.json`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      showToast("已下載 JSON 備份檔！");
    });

    // 還原 JSON 備份檔
    document.getElementById("fileImportBackup").addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const imported = JSON.parse(event.target.result);
          if (imported.medicalCare && imported.household && imported.transfers) {
            appData = imported;
            saveData();
            showToast("成功還原備份資料！");
            backupModal.classList.remove("open");
            renderDashboard();
            renderMedicalList();
            renderHouseholdList();
            renderTransferList();
          } else {
            alert("備份檔格式不符，請確認是本系統所匯出的 JSON 檔。");
          }
        } catch (err) {
          alert("讀取檔案失敗：" + err.message);
        }
      };
      reader.readAsText(file);
    });

    // 回復初始 Excel 資料
    document.getElementById("btnResetData").addEventListener("click", () => {
      if (confirm("注意：確定要將資料回復為 Excel 初始狀態嗎？您自訂新增的項目將會被重設。")) {
        if (window.INITIAL_DATA) {
          appData = JSON.parse(JSON.stringify(window.INITIAL_DATA));
          saveData();
          showToast("已回復為初始 Excel 資料！");
          backupModal.classList.remove("open");
          renderDashboard();
          renderMedicalList();
          renderHouseholdList();
          renderTransferList();
        }
      }
    });
  }

  // 程式進入點
  function init() {
    loadData();
    setupNavigation();
    setupFormToggles();
    setupFilters();
    setupForms();
    setupBackupAndExport();

    renderDashboard();
    renderMedicalList();
    renderHouseholdList();
    renderTransferList();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
