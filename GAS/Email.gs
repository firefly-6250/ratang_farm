function showUploadDialog() {
  const html = HtmlService.createTemplateFromFile('uploadexcel')
    .evaluate()
    .setWidth(450)
    .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, '黑貓託運結果上傳');
}

/**
 * 將前端傳來的資料寫入 Google Sheet
 * @param {Array[]} data 2D 陣列資料
 */
function processUploadedData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "黑貓寄件明細";
    let sheet = ss.getSheetByName(sheetName);

    // 如果工作表不存在，則建立它
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      // 每次上傳前先清空工作表所有內容與格式
      sheet.clear();
    }

    if (data && data.length > 0) {
      // 尋找「託運單號」所在的起始行（過濾掉檔案前幾行的標題或空白）
      let headerIndex = -1;
      for (let i = 0; i < data.length; i++) {
        // 檢查該列是否包含關鍵標題
        if (data[i] && data[i].indexOf("託運單號") !== -1) {
          headerIndex = i;
          break;
        }
      }

      if (headerIndex === -1) {
        throw new Error("找不到正確的標題行（檔案標題需包含：託運單號）");
      }

      // 取得從標題列開始的資料
      let actualData = data.slice(headerIndex);

      // 捨去最後兩行資料 (通常是報表末端的總計、頁碼或空白行)
      if (actualData.length > 2) {
        actualData = actualData.slice(0, -2);
      } else {
        // 如果資料不足兩行（只有標題），則不進行捨去或清空處理
        throw new Error("檔案資料筆數過少，無法執行捨去最後兩行的操作");
      }

      if (actualData.length > 0) {
        // 因為已經清空工作表，所以直接從第 1 列第 1 欄開始寫入
        sheet.getRange(1, 1, actualData.length, actualData[0].length).setValues(actualData);

        // 凍結第一列標題，方便管理者查看
        sheet.setFrozenRows(1);

        return { success: true, count: actualData.length - 1 }; // 回傳扣除標題後的資料筆數
      } else {
        return { success: true, count: 0, message: "解析完成，但沒有可寫入的資料內容" };
      }
    }

    return { success: false, message: "上傳的檔案不包含有效資料" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


function ShippedNotify() {
  // var ui = SpreadsheetApp.getUi();
  // var response = ui.alert('確認發送通知？', '是否確認使用配送明細發送出貨通知？', ui.ButtonSet.YES_NO);

  // if (response == ui.Button.NO) {
  //   return;
  // }

  // ================= ⚙️ 設定區 =================
  var sheepnumColIndex = getColumnNumberByHeader("sheepnum");
  var statusColIndex = getColumnNumberByHeader("status");
  var colIndex_Email = getColumnNumberByHeader("email"); // 欄位 N 是第 14 欄
  var colIndex_Phone = getColumnNumberByHeader("tel"); // 欄位 K 是第 11 欄
  var colIndex_SheepSearil = 2;
  var targetStatus = "已出貨";
  // ===========================================

  var shippingSheetName = "黑貓寄件明細";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var shippingSheet = ss.getSheetByName(shippingSheetName);
  var orderSheet = ss.getSheetByName(sourceSheetName);

  if (!shippingSheet || !orderSheet) {
    //SpreadsheetApp.getUi().alert("❌ 錯誤：找不到工作表，請檢查名稱。");
    console.error("錯誤：找不到工作表，請檢查名稱。");
    return;
  }

  var lastRow = shippingSheet.getLastRow();
  if (lastRow < 2) {
    //SpreadsheetApp.getUi().alert("⚠️ 出貨單沒有資料。");
    console.log("出貨單沒有資料。");
    return;
  }

  // 取得(託運單號) 和 (對應的列號/訂單編號)
  var rowNumbers = shippingSheet.getRange(2, 1, lastRow - 1, colIndex_SheepSearil).getValues();

  // 逐筆處理
  for (var i = 0; i < rowNumbers.length; i++) {
    var sheepNum = rowNumbers[i][0]; // 這裡是 A 欄：託運單號
    var rowNum = rowNumbers[i][1]; // 這裡是 B 欄：訂單編號

    // 檢查 rowNum 是否為有效數字
    if (rowNum && !isNaN(rowNum)) {
      try {
        // 取得該列的 Email 和電話
        var email = orderSheet.getRange(rowNum, colIndex_Email).getValue();
        var phone = orderSheet.getRange(rowNum, colIndex_Phone).getValue();

        // 更新狀態欄位填入託運單號
        var sheepnumCell = orderSheet.getRange(rowNum, sheepnumColIndex);
        var url = "https://www.t-cat.com.tw/Inquire/TraceDetail.aspx?BillID=" + sheepNum;
        sheepnumCell.setFormula('=HYPERLINK("' + url + '", "' + sheepNum + '")'); // 格式為：=HYPERLINK("網址", "顯示的文字")

        var statusCell = orderSheet.getRange(rowNum, statusColIndex);
        var currentStatus = String(statusCell.getValue());
        if (currentStatus.indexOf(targetStatus) === -1) {
          var newStatus = (currentStatus === "") ? targetStatus : currentStatus + "," + targetStatus;
          statusCell.setValue(newStatus);
        }

        // --- 發送通知邏輯 ---
        if (email && String(email).indexOf("@") !== -1) {
          // 情況 A：有 Email，發送郵件
          sendSheepEmail(email, sheepNum);
        } else {
          // 情況 B：沒有 Email (或格式錯誤)，改發簡訊
          if (phone) {
            // 處理手機號碼格式 (移除 "-" 和空白)
            var cleanPhone = String(phone).replace(/-/g, "").replace(/\s/g, "").trim();
            var message = "【訂單出貨】喇當大叔的果園已出貨給您,黑貓託運單號:" + sheepNum + "，請留意物流通知。感謝您的支持！";
            sendTwSms(cleanPhone, message);
          } else {
            console.log(">> 錯誤：列號 " + rowNum + " 沒有 Email 也沒有手機號碼");
          }
        }
      } catch (e) {
        console.error("處理列號 " + rowNum + " 時發生錯誤: " + e.message);
      }
    }
  }

  shippingSheet.clear(); //清空內容，避免重複發送。
  var today = new Date();
  var formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
  shippingSheet.getRange("A1").setValue(formattedDate);
  shippingSheet.getRange("B1").setValue("已完成出貨通知發送");

  //ui.alert("✅ 處理完成！");
}

function sendSheepEmail(targetEmail, sheep_Number) {
  try {
    // 1. 清理字串：移除 Email 內的所有空格，以及單號的前後空格
    var cleanEmail = targetEmail.toString().replace(/\s/g, "");
    var cleanNumber = sheep_Number.toString().trim();

    // 2. 組合郵件內容
    var subject = "【訂單出貨】喇當大叔的果園已出貨給您";
    var url = "https://www.t-cat.com.tw/Inquire/TraceDetail.aspx?BillID=" + cleanNumber;
    var body = "親愛的顧客您好：\n\n" +
      "您的訂單已經出貨，請留意物流通知。\n" +
      "下列網址可查詢黑貓託運進度：\n" + url +
      "\n\n感謝您的支持！";

    // 3. 執行發送
    GmailApp.sendEmail(cleanEmail, subject, body);
    console.log(">> 出貨Email 發送成功: " + cleanEmail + " (單號: " + cleanNumber + ")");

  } catch (err) {
    console.error(">> 出貨Email 發送失敗 (單號: " + sheep_Number + "): " + err.message);
  }
}


/**
 * 發送簡訊 (使用 TwSMS API 4.0 JSON版)
 *
 * @param {string} username TwSMS 帳號
 * @param {string} password TwSMS 密碼
 * @param {string} mobile   接收手機號碼 (台灣門號格式: 09xxxxxxxx)
 * @param {string} message  簡訊內容
 * @return {Object}         API 回傳的 JSON 物件 (包含 code, text, msgid 等)
 */
function sendTwSms(mobile, message) {
  // API URL (建議使用 HTTPS)
  var apiUrl = "https://api.twsms.com/json/sms_send.php";

  // 準備參數
  // 根據手冊，message 需要 urlencode，但在 GAS 的 payload 中通常會自動處理，
  // 若使用 GET 方法拼接 URL 則必須手動 encodeURIComponent。
  // 這裡使用 POST 方法比較乾淨。
  message = message + "(簡訊商:twsms)";
  var payload = {
    "username": "mamaratang",
    "password": "2lclgeu07e6chj3",
    "mobile": mobile,
    "message": message, // GAS UrlFetchApp 會自動處理 form-data 的編碼
    "drurl": "",        // (選填) 回傳狀態網址
    "mo": "N"           // (選填) 是否顯示雙向門號
  };

  // 設定 HTTP 請求選項
  var options = {
    "method": "post",
    "payload": payload,
    "muteHttpExceptions": true // 避免 HTTP 錯誤導致程式中斷
  };

  try {
    // 發送請求
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseText = response.getContentText();

    // 解析 JSON 回傳值
    var jsonResponse = JSON.parse(responseText);

    // 檢查結果
    if (jsonResponse.code === "00000") {
      Logger.log("簡訊發送成功！MsgID: " + jsonResponse.msgid + ", 號碼:" + mobile);
    } else {
      Logger.log("簡訊發送失敗。錯誤代碼: " + jsonResponse.code + ", 訊息: " + jsonResponse.text);
    }

    return jsonResponse;

  } catch (e) {
    Logger.log("發生錯誤: " + e.toString());
    return { "code": "99999", "text": "System Error: " + e.toString() };
  }
}

/**
 * 測試發送功能的範例函式
 */
function testSendSms() {
  var targetMobile = "0921709483";  // 請替換接收手機
  var myMessage = "【訂單出貨】喇當大叔的果園已出貨給您，請留意物流通知。感謝您的支持！\n換行測試"; // 支援 \n 換行

  var result = sendTwSms(targetMobile, myMessage);

  console.log(result);
}


// ---------------------------------------------------
// 📱 發送簡訊的功能 (需要串接付費 API)
// ---------------------------------------------------
function sendMitakeSMS(phoneNumber, message) {
  var apiUrl = "https://smsapi.mitake.com.tw/api/mtk/SmSend?&CharsetURL=UTF-8";
  var payload = {
    "username": "01014220A",
    "password": "",
    "dstaddr": phoneNumber,
    "smbody": message
  };

  var options = {
    "method": "post",
    "payload": payload
  };

  try {
    UrlFetchApp.fetch(apiUrl, options); // 解開註解才會真的發送
    console.log(">> 簡訊 API 呼叫成功，電話:" + phoneNumber + "，訊息:" + message);
  } catch (e) {
    console.error(">> 簡訊發送失敗: " + e.message);
  }
}

function test() {
  sendOrderEmail_HTML("yabung117@gmail.com", "ORD1779007059686", "plum")
  //sendOrderEmail_HTML("yabung117@gmail.com", "ORD1779007059686", "peach")
  //sendOrderEmail_HTML("farcheng132@gmail.com", "ORD1779007059686", "")
}

/**
 * 訂單成立通知信
 * @param {string} targetEmail 收件者 Email
 * @param {string} orderId     訂單編號
 * @param {string} orderType   'plum'＝李子、'peach'＝水蜜桃，其他（空值）＝香菇；
 *                             舊版傳入 truthy 布林值視同 'plum'
 */
function sendOrderEmail_HTML(targetEmail, orderId, orderType) {
  try {
    // 1. 空白檢查：檢查變數是否存在
    if (!targetEmail) {
      console.warn(">> 略過發送：第 " + orderId + " 行的 Email 為空值。");
      return;
    }

    // 2. 空格清理：移除字串中所有的空格、換行或縮排
    var cleanEmail = targetEmail.toString().replace(/\s/g, "");

    // 3. 清理後檢查：確保清理完後不是空字串
    if (cleanEmail === "") {
      console.warn(">> 略過發送：第 " + orderId + " 行的 Email 清理空格後無有效內容。");
      return;
    }

    var orderqueryurl = "https://ratang.pse.is/8mr8lr";
    var logoUrl = "https://i.meee.com.tw/qbs7iFL.jpg"; // ← 請改成你的Logo圖片網址

    var subject = "【訂單通知】喇當大叔的果園已收到您的訂單(訂單編號:" + orderId + ")";
    var htmlBody_plum = `
<div style="margin:0; padding:0; background-color:#f8f5f0;">
  <div style="max-width:620px; margin:60px auto; padding:0 24px;">

    <div style="
      background-color:#ffffff;
      border-radius:18px;
      padding:48px 40px;
      box-shadow:0 12px 32px rgba(0,0,0,0.06);
      font-family: 'Microsoft JhengHei', Arial, sans-serif;
      line-height:1.9;
      color:#2b2b2b;
    ">

      <div style="text-align:center; margin-bottom:40px;">
        <img src="${logoUrl}" style="max-width:140px;">
      </div>

      <p style="font-size:16px;">
        親愛的朋友，您好：
      </p>

      <p>
        感謝您的訂購，我們已經確實收到您的訂單了！
      </p>

      <p>
        我們的李子，歷經高海拔日夜溫差的嚴苛考驗，也堅持不噴灑化學藥劑的野放種植，全程仰賴人工除草與雙手採摘，每一顆結實纍纍的果實都來之不易。
      </p>

      <p>
        為了讓這份美好完好如初地送達，我們特意在果實七、八分熟時採下，讓這批後熟的水果隨著時間慢慢呼吸、轉化，這份香氣與甜度也會在最佳狀態抵達您手中。
      </p>

      <p>
        果園目前正依序安排採收，出貨當天我們會再發送通知給您，請您安心期待。
      </p>

      <p>
        願當您開箱的那一刻，能迎來撲鼻的山林果香，並真切感受到來自產地的真實、純粹與用心。
      </p>

      <p>
        再次感謝您的等待與支持。
      </p>

      <p style="margin-top:40px;">
        謹致<br>
        <strong style="font-size:18px; letter-spacing:2px;">喇當大叔的果園</strong>
      </p>

      <hr style="border: none; border-top: 1px solid #eee; margin: 32px 0;">

      <div style="font-size: 14px; color: #666666; line-height: 1.6;">
        <p style="margin: 0 0 8px 0;">
        <strong>聯絡資訊：</strong>
        </p>
        <p style="margin: 0;">
          此信件為系統自動發送，若有任何疑問，請私訊 
          <a href="https://www.facebook.com/RatangsFruitFarm" target="_blank" style="color: #1877f2; text-decoration: underline; font-weight: bold;">
            「喇當大叔的果園」臉書粉絲專頁
          </a>，並請於私訊時提供<strong>收件人姓名</strong>，謝謝您。
        </p>
      </div>

    </div>

  </div>
</div>
`;

    var htmlBody_peach = `
<div style="margin:0; padding:0; background-color:#f8f5f0;">
  <div style="max-width:620px; margin:60px auto; padding:0 24px;">

    <div style="
      background-color:#ffffff;
      border-radius:18px;
      padding:48px 40px;
      box-shadow:0 12px 32px rgba(0,0,0,0.06);
      font-family: 'Microsoft JhengHei', Arial, sans-serif;
      line-height:1.9;
      color:#2b2b2b;
    ">

      <div style="text-align:center; margin-bottom:40px;">
        <img src="${logoUrl}" style="max-width:140px;">
      </div>

      <p style="font-size:16px;">
        親愛的朋友，您好：
      </p>

      <p>
        感謝您的訂購，我們已經確實收到您的訂單了！
      </p>

      <p>
        我們的水蜜桃，生長在海拔約 1700 公尺的合歡山西側山腰——馬烈霸部落。日夜溫差極大的環境、純淨的山泉水澆灌，加上謹守減少用藥原則的細心照顧，讓每一顆果實都能累積最飽滿的甜分與濃郁的果香。
      </p>

      <p>
        水蜜桃十分嬌貴，我們會依果實的熟度採收，並按照訂購順序陸續出貨；出貨當天我們會再發送通知給您，請您安心期待。
      </p>

      <p>
        收到後建議先冷藏保存並盡早享用，最能品嚐到水蜜桃細緻多汁、香甜純粹的風味。
      </p>

      <p>
        再次感謝您的等待與支持。
      </p>

      <p style="margin-top:40px;">
        謹致<br>
        <strong style="font-size:18px; letter-spacing:2px;">喇當大叔的果園</strong>
      </p>

      <hr style="border: none; border-top: 1px solid #eee; margin: 32px 0;">

      <div style="font-size: 14px; color: #666666; line-height: 1.6;">
        <p style="margin: 0 0 8px 0;">
        <strong>聯絡資訊：</strong>
        </p>
        <p style="margin: 0;">
          此信件為系統自動發送，若有任何疑問，請私訊
          <a href="https://www.facebook.com/RatangsFruitFarm" target="_blank" style="color: #1877f2; text-decoration: underline; font-weight: bold;">
            「喇當大叔的果園」臉書粉絲專頁
          </a>，並請於私訊時提供<strong>收件人姓名</strong>，謝謝您。
        </p>
      </div>

    </div>

  </div>
</div>
`;

    var htmlBody = `
<div style="margin:0; padding:0; background-color:#f8f5f0;">
  <div style="max-width:620px; margin:60px auto; padding:0 24px;">

    <div style="
      background-color:#ffffff;
      border-radius:18px;
      padding:48px 40px;
      box-shadow:0 12px 32px rgba(0,0,0,0.06);
      font-family: 'Microsoft JhengHei', Arial, sans-serif;
      line-height:1.9;
      color:#2b2b2b;
    ">

      <div style="text-align:center; margin-bottom:40px;">
        <img src="${logoUrl}" style="max-width:140px;">
      </div>

      <p style="font-size:16px;">
        親愛的顧客您好：
      </p>

      <p>
        我們在最完美的時機採收，歷經多道手工繁複的洗淨、日曬與焙火，只為鎖住更濃郁的菇香，留下大自然賜予的厚實與甘甜。
      </p>

      <p>
        願當您開箱時，能聞到撲鼻而來的溫暖香氣，<br>
        感受到來自產地的真實、純粹與用心。
      </p>

      <p>
        不論是用來熬湯、入菜，希望這份厚實的滋味能溫暖您的餐桌。
      </p>

      <p style="margin-top:40px;">
        謹致<br>
        <strong style="font-size:18px; letter-spacing:2px;">喇當大叔的果園X紅香部落椴木香菇</strong>
      </p>

      <hr style="border: none; border-top: 1px solid #eee; margin: 32px 0;">

      <div style="font-size: 14px; color: #666666; line-height: 1.6;">
        <p style="margin: 0 0 8px 0;">
        <strong>聯絡資訊：</strong>
        </p>
        <p style="margin: 0;">
          此信件為系統自動發送，若有任何疑問，請私訊 
          <a href="https://www.facebook.com/RatangsFruitFarm" target="_blank" style="color: #1877f2; text-decoration: underline; font-weight: bold;">
            「喇當大叔的果園」臉書粉絲專頁
          </a>，並請於私訊時提供<strong>收件人姓名</strong>，謝謝您。
        </p>
      </div>

    </div>

  </div>
</div>
`;

    // 依訂單類型選擇信件模板（預設：香菇；舊版 truthy 布林值視同李子）
    var htmlBodyToSend = htmlBody;
    if (orderType === 'peach') {
      htmlBodyToSend = htmlBody_peach;
    } else if (orderType === 'plum' || (orderType && orderType !== '')) {
      htmlBodyToSend = htmlBody_plum;
    }

    GmailApp.sendEmail(cleanEmail, subject, "",
      {
        htmlBody: htmlBodyToSend,
        name: "喇當大叔的果園",
        charset: "UTF-8"
      });

    console.log(">> Email 發送成功: " + cleanEmail + " (orderId: " + orderId + ")");
  } catch (err) {
    // 在錯誤紀錄中加入 orderRow，方便回頭檢查工作表哪一筆出錯
    console.error(">> 訂單Email 發送失敗 (orderId: " + orderId + "): " + err.message);
  }
}


