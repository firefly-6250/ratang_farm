/**
 * 試算表設定
 */
const SPREADSHEET_ID = '13s8MhC6LmZSA0Brbal7bigI_n4QaJuqWMiSg6BfL1kk';
const ORDER_SHEET = '訂單';

/**
 * 主函式：由 GAS 時間型排程每小時呼叫一次。
 *
 * 防重複機制：發送通知後，將 sheep_id 欄的純文字改寫為 HYPERLINK 公式。
 * 下次掃描時，若 sheep_id 欄已含 HYPERLINK 公式則視為已通知，直接跳過。
 *
 * 列結構假設：第 1 列為英文標題，第 2 列為中文說明，第 3 列起為訂單資料。
 */
function runHourlyCheck() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_SHEET);
  if (!sheet) {
    Logger.log("找不到工作表：" + ORDER_SHEET);
    return;
  }

  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastColumn === 0 || lastRow < 3) {
    Logger.log("工作表無資料列，跳過。");
    return;
  }

  // 讀取標題列，定位各欄位
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var statusCol  = headers.indexOf("status")   + 1;
  var sheepIdCol = headers.indexOf("sheep_id") + 1;
  var emailCol   = headers.indexOf("email")    + 1;
  var telCol     = headers.indexOf("tel")      + 1;

  if (statusCol === 0) {
    Logger.log("找不到 'status' 欄位，終止。");
    return;
  }
  if (sheepIdCol === 0) {
    Logger.log("找不到 'sheep_id' 欄位，終止。");
    return;
  }

  // 一次讀取所有資料列（第 3 列起）的值與公式
  var dataStartRow = 3;
  var numRows = lastRow - dataStartRow + 1;
  if (numRows <= 0) {
    Logger.log("無訂單資料。");
    return;
  }

  var dataRange = sheet.getRange(dataStartRow, 1, numRows, lastColumn);
  var data     = dataRange.getValues();
  var formulas = dataRange.getFormulas();
  var sentCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var status = row[statusCol - 1] != null ? row[statusCol - 1].toString() : "";

    // sheep_id 欄已含 HYPERLINK 公式 → 已通知過，跳過
    var sheepIdFormula = formulas[i][sheepIdCol - 1];
    if (sheepIdFormula.toUpperCase().indexOf("HYPERLINK") !== -1) continue;

    // 狀態包含「已出貨」才處理
    if (!status.includes("已出貨")) continue;

    var orderId = row[sheepIdCol - 1] !== "" ? row[sheepIdCol - 1].toString() : "無訂單編號";
    var email   = (emailCol > 0) ? row[emailCol - 1].toString() : "";
    var phone   = (telCol   > 0) ? row[telCol   - 1].toString() : "";

    handleOrderNotification({ email: email, phone: phone }, orderId);

    // 發送後將 sheep_id 欄改寫為 HYPERLINK 公式，作為「已通知」的標記
    if (orderId !== "無訂單編號") {
      var url = generatePostEqueryUrl(orderId);
      var sheetRow = dataStartRow + i;
      sheet.getRange(sheetRow, sheepIdCol)
           .setFormula('=HYPERLINK("' + url + '","' + orderId + '")');
    }
    sentCount++;
  }

  Logger.log(sentCount > 0
    ? "本次共發送 " + sentCount + " 筆通知。"
    : "本次無需發送通知。");
}

/**
 * 核心判斷邏輯：有效 Email 發信，否則發簡訊
 */
function handleOrderNotification(payload, orderId) {
  if (isValidEmail(payload.email)) {
    sendSheepEmail(payload.email, orderId);
  } else {
    sendOrderSMS(payload.phone, orderId);
  }
}

/**
 * 驗證 Email 格式
 */
function isValidEmail(email) {
  if (!email) return false;
  var emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email.toString().trim());
}

/**
 * 發送 HTML 信件
 */
function sendSheepEmail(targetEmail, orderId) {
  var cleanEmail = "";
  var cleanNumber = "";
  try {
    cleanEmail  = targetEmail.toString().replace(/\s/g, "");
    cleanNumber = orderId.toString().trim();

    var subject = "【訂單出貨】喇當大叔的果園已出貨給您(郵件號碼:" + cleanNumber + ")";
    var url = generatePostEqueryUrl(cleanNumber);
    var body = "親愛的顧客您好：\n\n" +
      "您的訂單已經出貨，請留意物流通知。\n" +
      "下列網址可查詢託運進度：\n" + url + "\n" +
      "郵件號碼:" + cleanNumber + "\n\n感謝您的支持！";

    GmailApp.sendEmail(cleanEmail, subject, body);
    console.log(">> 出貨Email 發送成功: " + cleanEmail + " (單號: " + cleanNumber + ")");
  } catch (err) {
    console.error(">> 出貨Email 發送失敗 (單號: " + (cleanNumber || orderId) + "): " + err.message);
  }
}

/**
 * 產生 equery API 查詢連結（較簡潔）
 * @param {string} mailno - 郵件號碼
 * @param {number} type - 查詢類型，預設 1（國內）
 * @returns {string} 完整查詢 URL
 */
function generatePostEqueryUrl(mailno, type) {
  type = type || 1;
  var mailnoB64 = Utilities.base64Encode(mailno);
  return 'http://postserv.post.gov.tw/pstmail/equery?type=' + type + '&mailno=' + encodeURIComponent(mailnoB64);
}

/**
 * 發送簡訊
 */
function sendOrderSMS(phone, orderId) {
  console.log(">> 正在發送簡訊給: " + phone + " (單號: " + orderId + ")");
  var message = "【訂單出貨】喇當大叔的果園已出貨給您(郵件號碼:" + orderId + ")。請留意物流通知。";
  sendTwSms(phone, message);
}

/**
 * 測試用：直接呼叫 runHourlyCheck，可在 GAS 編輯器手動點「執行」驗證。
 * 若要重測已通知的列，先手動把該列 sheep_id 欄清回純文字單號，再執行此函式。
 */
function test_runHourlyCheck() {
  //runHourlyCheck();
  var url = generatePostEqueryUrl("72047890010770411002");
  console.log(url);

}
