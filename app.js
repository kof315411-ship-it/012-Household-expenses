// 家庭照顧與家用花費紀錄 APP 核心邏輯
(function () {
  "use strict";

  const STORAGE_KEY = "FAMILY_EXPENSE_DATA_V1";
  const SPREADSHEET_ID = "1wut51zEYI7Ij0aBbCx-b7fpEA_n0lhdMht7PtEjnGHY";

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

  // 工具函式：格式化千分位貨幣
  function formatMoney(num) {
    if (num === null || num === undefined || isNaN(num)) return "NT$ 0";
    return "NT$ " + Math.round(num).toLocaleString("en-US");
  }

  // 工具函式：解析各種日期字串為 YYYY-MM-DD
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

  // 工具函式：取得今日日期 (YYYY-MM-DD)
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
    }, 2800);
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
    const lastSync = localStorage.getItem("LAST_CLOUD_SYNC");
    if (badge) {
      if (lastSync) {
        badge.textContent = `已同步 (${lastSync})`;
      } else {
        badge.textContent = `已連結`;
      }
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
    // 就醫表單
    const btnMed = document.getElementById("btnToggleMedForm");
    const cardMed = document.getElementById("formMedCard");
    const btnCancelMed = document.getElementById("btnCancelMed");
    btnMed.addEventListener("click", () => {
      cardMed.classList.toggle("open");
    });
    btnCancelMed.addEventListener("click", () => {
      cardMed.classList.remove("open");
    });

    // 家用表單
    const btnHouse = document.getElementById("btnToggleHouseForm");
    const cardHouse = document.getElementById("formHouseCard");
    const btnCancelHouse = document.getElementById("btnCancelHouse");
    btnHouse.addEventListener("click", () => {
      cardHouse.classList.toggle("open");
    });
    btnCancelHouse.addEventListener("click", () => {
      cardHouse.classList.remove("open");
    });

    // 老哥轉帳表單
    const btnTransfer = document.getElementById("btnToggleTransferForm");
    const cardTransfer = document.getElementById("formTransferCard");
    const btnCancelTransfer = document.getElementById("btnCancelTransfer");
    btnTransfer.addEventListener("click", () => {
      cardTransfer.classList.toggle("open");
    });
    btnCancelTransfer.addEventListener("click", () => {
      cardTransfer.classList.remove("open");
    });

    // 預設日期為今日
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
      showToast("已刪除紀錄");
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

  // 關閉編輯 Modal
  function closeEditModal() {
    document.getElementById("editModal").classList.remove("open");
  }

  // 綁定篩選器監聽
  function setupFilters() {
    // 就醫搜尋
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

    // 家用搜尋
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

    // 老哥轉帳搜尋
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
    // 1. 新增就醫紀錄
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
      showToast("已成功新增就醫花費紀錄！");
      document.getElementById("medAmount").value = "";
      document.getElementById("medNote").value = "";
      document.getElementById("formMedCard").classList.remove("open");
      renderMedicalList();
      renderDashboard();
    });

    // 2. 新增家用紀錄
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
      showToast("已成功新增家用繳款紀錄！");
      document.getElementById("houseAmount").value = "";
      document.getElementById("houseNote").value = "";
      document.getElementById("formHouseCard").classList.remove("open");
      renderHouseholdList();
      renderDashboard();
    });

    // 3. 新增老哥轉帳紀錄
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
      showToast("已成功新增老哥轉帳紀錄！");
      document.getElementById("transferAmount").value = "";
      document.getElementById("transferNote").value = "";
      document.getElementById("formTransferCard").classList.remove("open");
      renderTransferList();
      renderDashboard();
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

    // 編輯 Modal 取消按鈕
    document.getElementById("btnCancelEdit").addEventListener("click", closeEditModal);
    document.getElementById("btnCloseEditModal").addEventListener("click", closeEditModal);
  }

  // Google 雲端試算表 JSONP 擷取器
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

  // 執行全量雲端試算表同步
  async function syncFromGoogleSheets() {
    const btnSync = document.getElementById("btnSyncCloud");
    const btnModalSync = document.getElementById("btnModalSyncCloud");
    const syncIcon = btnSync ? btnSync.querySelector(".sync-icon") : null;

    if (syncIcon) syncIcon.classList.add("spinning");
    if (btnSync) btnSync.disabled = true;
    if (btnModalSync) btnModalSync.disabled = true;

    showToast("⏳ 正在從 Google 雲端試算表下載最新資料...");

    try {
      // 平行拉取 3 個工作表
      const [rowsMed, rowsHouse, rowsTransfer] = await Promise.all([
        fetchSheetDataJSONP("就醫、照顧花費記帳"),
        fetchSheetDataJSONP("家用花費、繳款紀錄"),
        fetchSheetDataJSONP("老哥轉帳記錄")
      ]);

      // 1. 解析就醫照顧
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

      // 2. 解析家用花費
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
        if (rawItem.includes("台電")) {
          itemCat = "台電";
        } else if (rawItem.includes("中華電信")) {
          itemCat = "中華電信";
        } else if (rawItem.includes("瓦斯")) {
          itemCat = "瓦斯";
        } else if (rawItem.includes("北水") || rawItem.includes("水")) {
          itemCat = "北水";
        } else if (rawItem.includes("轉帳")) {
          itemCat = "房貸轉帳";
        } else if (rawItem.includes("墓園管理費")) {
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

      // 3. 解析老哥轉帳
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

      // 更新全域狀態
      appData = {
        medicalCare: newMed,
        household: newHouse,
        transfers: newTransfer
      };

      saveData();
      const nowStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      localStorage.setItem("LAST_CLOUD_SYNC", nowStr);
      updateSyncBadge();

      // 重新繪製介面
      renderDashboard();
      renderMedicalList();
      renderHouseholdList();
      renderTransferList();

      showToast("🎉 已成功從 Google 雲端試算表同步最新資料！");
    } catch (err) {
      console.error("雲端同步失敗:", err);
      showToast("❌ 雲端同步失敗：" + err.message);
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
      backupModal.classList.add("open");
    });
    document.getElementById("btnCloseBackupModal").addEventListener("click", () => {
      backupModal.classList.remove("open");
    });

    // 雲端同步按鈕綁定
    const btnSyncCloud = document.getElementById("btnSyncCloud");
    if (btnSyncCloud) btnSyncCloud.addEventListener("click", syncFromGoogleSheets);

    const btnModalSyncCloud = document.getElementById("btnModalSyncCloud");
    if (btnModalSyncCloud) btnModalSyncCloud.addEventListener("click", syncFromGoogleSheets);

    // 匯出 CSV (支援 Excel UTF-8 with BOM)
    function exportToCSV() {
      let csvContent = "\uFEFF"; // UTF-8 BOM
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
