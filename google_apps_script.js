/**
 * ========================================================
 * Google 試算表自動同步後端腳本 (Google Apps Script)
 * 試算表：家庭與老爸照顧花費＆照顧紀錄
 * 網址：https://docs.google.com/spreadsheets/d/1wut51zEYI7Ij0aBbCx-b7fpEA_n0lhdMht7PtEjnGHY/edit
 * ========================================================
 * 
 * 【部署說明】：
 * 1. 在 Google 試算表中，點擊上方選單「擴充功能」 -> 「Apps Script」
 * 2. 刪除原有程式碼，將本檔案所有內容完整貼上
 * 3. 點擊右上角「部署」 -> 「管理部署」或「新部署」
 *    - 若選「管理部署」：點右上角鉛筆圖示，版本選「新版本」，點「部署」
 *    - 若選「新部署」：
 *      - 種類：網頁應用程式 (Web App)
 *      - 執行身分：我 (您的 Google 帳號)
 *      - 誰可以存取：任何人 (Anyone)  <--- 最重要！
 * 4. 點擊「部署」，授權存取，並複製產生的「網頁應用程式網址」
 * 5. 回到手機/電腦網頁工具的「⚙️ 雲端設定」，將網址貼入「雲端同步網址」儲存即可！
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

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
    } else if (action === "update" || action === "edit") {
      output = updateRecord(ss, category, data, params.oldData);
    } else if (action === "delete") {
      output = deleteRecord(ss, category, data);
    } else if (action === "batch_add") {
      output = batchAddRecords(ss, params.items || []);
    } else {
      output = { success: false, message: "未知的操作指令：" + action };
    }
  } catch (err) {
    output = { success: false, message: "執行錯誤：" + err.toString() };
  }

  // 支援 JSONP 與 JSON 回傳，避免瀏覽器跨域限制
  var callback = (e.parameter && e.parameter.callback);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(output) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 新增單筆紀錄至試算表
 */
function addRecord(ss, category, data) {
  var dateStr = formatDateToYYMMDD(data.date);
  var amount = Number(data.amount) || 0;
  var note = data.note || "";

  if (category === "medical") {
    // 工作表：就醫、照顧花費記帳
    var sheet = ss.getSheetByName("就醫、照顧花費記帳");
    if (!sheet) return { success: false, message: "找不到工作表「就醫、照顧花費記帳」" };

    // 推算分類 (照護 / 醫療 / 其他)
    var item = data.item || "其他";
    var type = "其他";
    if (/看診|住院/.test(item)) type = "醫療";
    else if (/長照|看護/.test(item)) type = "照護";

    // 檢查是否有老媽款項
    var momAmt = "";
    if (note.indexOf("老媽") !== -1) {
      momAmt = amount;
    }

    var targetRow = findFirstEmptyRow(sheet, 2, 4);

    // 寫入欄位：B:日期, C:項目, D:分類, E:金額, F:備註, G:老媽
    sheet.getRange(targetRow, 2).setValue(dateStr);
    sheet.getRange(targetRow, 3).setValue(item);
    sheet.getRange(targetRow, 4).setValue(type);
    sheet.getRange(targetRow, 5).setValue(amount);
    sheet.getRange(targetRow, 6).setValue(note);
    if (momAmt !== "") {
      sheet.getRange(targetRow, 7).setValue(momAmt);
    }

    return { success: true, message: "已成功新增至雲端「就醫、照顧花費記帳」第 " + targetRow + " 列！", row: targetRow };

  } else if (category === "household") {
    // 工作表：家用花費、繳款紀錄
    var sheet = ss.getSheetByName("家用花費、繳款紀錄");
    if (!sheet) return { success: false, message: "找不到工作表「家用花費、繳款紀錄」" };

    var item = data.item || "其他";
    var account = "";
    var type = "其他";

    if (item === "房貸轉帳") {
      item = "轉帳";
      account = "永豐";
      type = "房貸";
    } else if (/台電|北水|瓦斯/.test(item)) {
      type = "水電";
      account = "竑郵局";
    } else if (item === "中華電信") {
      type = "其他";
      account = "竑郵局";
    }

    var targetRow = findFirstEmptyRow(sheet, 2, 4);

    // 寫入欄位：B:日期, C:項目, D:帳戶, E:分類, F:金額, G:備註
    sheet.getRange(targetRow, 2).setValue(dateStr);
    sheet.getRange(targetRow, 3).setValue(item);
    sheet.getRange(targetRow, 4).setValue(account);
    sheet.getRange(targetRow, 5).setValue(type);
    sheet.getRange(targetRow, 6).setValue(amount);
    sheet.getRange(targetRow, 7).setValue(note);

    return { success: true, message: "已成功新增至雲端「家用花費、繳款紀錄」第 " + targetRow + " 列！", row: targetRow };

  } else if (category === "transfers") {
    // 工作表：老哥轉帳記錄
    var sheet = ss.getSheetByName("老哥轉帳記錄");
    if (!sheet) return { success: false, message: "找不到工作表「老哥轉帳記錄」" };

    var targetRow = findFirstEmptyRow(sheet, 2, 5);

    // 寫入欄位：B:日期, G:換匯後金額, H:備註
    sheet.getRange(targetRow, 2).setValue(dateStr);
    sheet.getRange(targetRow, 7).setValue(amount);
    sheet.getRange(targetRow, 8).setValue(note);

    return { success: true, message: "已成功新增至雲端「老哥轉帳記錄」第 " + targetRow + " 列！", row: targetRow };
  }

  return { success: false, message: "不支援的類別：" + category };
}

/**
 * 批次新增多筆紀錄
 */
function batchAddRecords(ss, items) {
  if (!items || !items.length) {
    return { success: false, message: "無待新增紀錄" };
  }
  var successCount = 0;
  for (var i = 0; i < items.length; i++) {
    var res = addRecord(ss, items[i].category, items[i].data);
    if (res && res.success) successCount++;
  }
  return { success: true, message: "批次同步完成！成功寫入 " + successCount + " 筆至雲端試算表。" };
}

/**
 * 尋找相符的列 (從後向前逆向掃描，支援 Date 物件比對與老媽金額合併比對)
 */
function findMatchingRow(sheet, category, data, oldData) {
  var targetData = oldData || data;
  var targetDate = normalizeDateYYMMDD(targetData.date);
  var targetAmt = Math.round(Number(targetData.amount) || 0);
  var targetItem = String(targetData.item || "").trim();

  var maxRows = sheet.getLastRow();
  var startRow = category === "transfers" ? 5 : 4;
  if (maxRows < startRow) return -1;

  var amtCol = category === "medical" ? 5 : (category === "household" ? 6 : 7);
  var itemCol = (category === "medical" || category === "household") ? 3 : -1;
  var dataValues = sheet.getRange(startRow, 1, maxRows - startRow + 1, 8).getValues();

  var candidateRow = -1;

  for (var i = dataValues.length - 1; i >= 0; i--) {
    var row = dataValues[i];
    var rowDate = normalizeDateYYMMDD(row[1]);
    var rowAmt = Math.round(Number(row[amtCol - 1]) || 0);
    var rowItem = itemCol > 0 ? String(row[itemCol - 1] || "").trim() : "";

    if (rowDate !== targetDate) continue;

    var amtMatch = (rowAmt === targetAmt);
    if (!amtMatch && category === "medical") {
      var rowMomAmt = Math.round(Number(row[6]) || 0); // Column G (老媽)
      if ((rowAmt + rowMomAmt === targetAmt) || (rowMomAmt === targetAmt)) {
        amtMatch = true;
      }
    }

    if (amtMatch) {
      var itemMatch = (itemCol <= 0 || !targetItem || !rowItem || rowItem === targetItem || rowItem.indexOf(targetItem) !== -1 || targetItem.indexOf(rowItem) !== -1);
      if (itemMatch) {
        return startRow + i;
      } else if (candidateRow === -1) {
        candidateRow = startRow + i;
      }
    }
  }

  return candidateRow;
}

/**
 * 修改試算表既有紀錄
 */
function updateRecord(ss, category, data, oldData) {
  var sheetName = category === "medical" ? "就醫、照顧花費記帳" : (category === "household" ? "家用花費、繳款紀錄" : "老哥轉帳記錄");
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: "找不到工作表「" + sheetName + "」" };

  var targetRow = findMatchingRow(sheet, category, data, oldData);
  if (targetRow === -1) {
    var addRes = addRecord(ss, category, data);
    return {
      success: true,
      message: "雲端試算表未找到舊列，已自動補增為第 " + (addRes.row || "新") + " 列！",
      row: addRes.row
    };
  }

  var newDate = formatDateToYYMMDD(data.date);
  var newAmt = Number(data.amount) || 0;
  var newNote = data.note || "";

  sheet.getRange(targetRow, 2).setValue(newDate);

  if (category === "medical") {
    var item = data.item || "其他";
    var type = "其他";
    if (/看診|住院/.test(item)) type = "醫療";
    else if (/長照|看護/.test(item)) type = "照護";

    sheet.getRange(targetRow, 3).setValue(item);
    sheet.getRange(targetRow, 4).setValue(type);
    sheet.getRange(targetRow, 5).setValue(newAmt);
    sheet.getRange(targetRow, 6).setValue(newNote);
    if (newNote.indexOf("老媽") !== -1) {
      sheet.getRange(targetRow, 7).setValue(newAmt);
    } else {
      sheet.getRange(targetRow, 7).setValue("");
    }
  } else if (category === "household") {
    var item = data.item || "其他";
    var account = "";
    var type = "其他";

    if (item === "房貸轉帳") {
      item = "轉帳";
      account = "永豐";
      type = "房貸";
    } else if (/台電|北水|瓦斯/.test(item)) {
      type = "水電";
      account = "竑郵局";
    } else if (item === "中華電信") {
      type = "其他";
      account = "竑郵局";
    }

    sheet.getRange(targetRow, 3).setValue(item);
    if (account) sheet.getRange(targetRow, 4).setValue(account);
    if (type) sheet.getRange(targetRow, 5).setValue(type);
    sheet.getRange(targetRow, 6).setValue(newAmt);
    sheet.getRange(targetRow, 7).setValue(newNote);
  } else if (category === "transfers") {
    sheet.getRange(targetRow, 7).setValue(newAmt);
    sheet.getRange(targetRow, 8).setValue(newNote);
  }

  return { success: true, message: "已成功更新雲端試算表第 " + targetRow + " 列！", row: targetRow };
}

/**
 * 刪除試算表紀錄
 */
function deleteRecord(ss, category, data) {
  var sheetName = category === "medical" ? "就醫、照顧花費記帳" : (category === "household" ? "家用花費、繳款紀錄" : "老哥轉帳記錄");
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: "找不到工作表「" + sheetName + "」" };

  var targetRow = findMatchingRow(sheet, category, data, null);
  if (targetRow === -1) {
    return { success: true, message: "雲端試算表未找到該筆紀錄 (可能已被刪除)，本機已刪除！" };
  }

  sheet.deleteRow(targetRow);
  return { success: true, message: "已從雲端試算表刪除第 " + targetRow + " 列！", row: targetRow };
}

/**
 * 輔助函式：標準化日期為 YYMMDD 字串格式 (如 "220831" 或 "260925")
 */
function normalizeDateYYMMDD(val) {
  if (val === null || val === undefined || val === "") return "";
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    var d = val.getDate();
    var yy = String(y).slice(-2);
    var mm = m < 10 ? "0" + m : "" + m;
    var dd = d < 10 ? "0" + d : "" + d;
    return yy + mm + dd;
  }
  var s = String(val).trim().split('.')[0].replace(/[-/]/g, "");
  if (s.length === 8) {
    return s.substring(2);
  } else if (s.length === 6) {
    return s;
  }
  var parsed = new Date(val);
  if (!isNaN(parsed.getTime())) {
    var y2 = parsed.getFullYear();
    var m2 = parsed.getMonth() + 1;
    var d2 = parsed.getDate();
    var yy2 = String(y2).slice(-2);
    var mm2 = m2 < 10 ? "0" + m2 : "" + m2;
    var dd2 = d2 < 10 ? "0" + d2 : "" + d2;
    return yy2 + mm2 + dd2;
  }
  return s;
}

/**
 * 輔助函式：將日期轉成試算表存儲格式 (數值 260925)
 */
function formatDateToYYMMDD(val) {
  var norm = normalizeDateYYMMDD(val);
  var n = Number(norm);
  return isNaN(n) ? norm : n;
}

/**
 * 尋找指定欄位從 startRow 開始的第一個空白列
 */
function findFirstEmptyRow(sheet, colIndex, startRow) {
  var maxRows = sheet.getMaxRows();
  var values = sheet.getRange(startRow, colIndex, maxRows - startRow + 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v === "" || v === null || v === undefined) {
      return startRow + i;
    }
  }
  sheet.appendRow([""]);
  return sheet.getLastRow();
}
