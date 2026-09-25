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
 * 3. 點擊右上角「部署」 -> 「新部署」
 * 4. 點選左側齒輪，選擇「網頁應用程式 (Web App)」
 *    - 說明：家庭記帳即時同步
 *    - 執行身分：我 (您的 Google 帳號)
 *    - 誰可以存取：任何人 (Anyone)  <--- 重要！這樣手機才能免登入直接同步
 * 5. 點擊「部署」，授權存取，並複製產生的「網頁應用程式網址」
 * 6. 回到手機/電腦網頁工具的「⚙️ 設定」，將網址貼入「雲端同步網址」儲存即可！
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
      momAmt = amount; // 若有註明老媽
    }

    // 尋找最後一筆有日期的列 (檢查 B 欄)
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
 * 輔助函式：將 YYYY-MM-DD 或 Date 轉成試算表使用的 YYMMDD (數值格式，如 260925)
 */
function formatDateToYYMMDD(val) {
  if (!val) {
    var now = new Date();
    var yy = String(now.getFullYear()).substring(2);
    var mm = String(now.getMonth() + 1).padStart(2, "0");
    var dd = String(now.getDate()).padStart(2, "0");
    return Number(yy + mm + dd);
  }
  var s = String(val).replace(/[-/]/g, "").trim();
  if (s.length === 8) {
    // 20260925 -> 260925
    return Number(s.substring(2));
  } else if (s.length === 6) {
    return Number(s);
  }
  return val;
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
  // 若現有列已滿，在下方新增一列
  sheet.appendRow([""]);
  return sheet.getLastRow();
}
