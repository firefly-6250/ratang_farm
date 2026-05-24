
function processPostOfficeDeposit() {
  // ================= ⚙️ 參數設定區 =================
  var labelName = '已處理';
  var logSheetName = '郵局匯款紀錄';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var logSheet = ss.getSheetByName(logSheetName);
  var orderSheet = ss.getSheetByName("訂單");
  var orderSheet2 = ss.getSheetByName("李子訂單");
  // ================= ⚙️ 參數設定區 =================

  if (!orderSheet) {
    console.log("錯誤：找不到訂單工作表");
    return;
  }

  function buildSheetInfo(sheet) {
    if (!sheet) return null;
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var getIdx = function(name) {
      var i = headerRow.indexOf(name);
      return i === -1 ? null : i + 1;
    };
    var colIdx_Last5 = getIdx("ending5");
    return {
      sheet: sheet,
      sheetId: sheet.getSheetId(),
      colIdx_Last5: colIdx_Last5,
      colIdx_Amount: getIdx("amount"),
      colIdx_NeedAmount: getIdx("shouldpay"),
      colIdx_Status: getIdx("status"),
      colIdx_Email: getIdx("email"),
      colIdx_Phone: getIdx("tel"),
      colIdx_Name: getIdx("name"),
      colLetter_Last5: sheet.getRange(1, colIdx_Last5).getA1Notation().replace(/[0-9]/g, '')
    };
  }

  var sheetInfo1 = buildSheetInfo(orderSheet);
  var sheetInfo2 = buildSheetInfo(orderSheet2);

  var headers = ['轉入日期', '轉入時間', '轉入金額', '轉出帳號', '帳號末五碼', '轉出行庫', '對帳狀態', '系統寫入時間', '客戶姓名'];
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
    logSheet.appendRow(headers);
  }

  var label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
  var query = 'label:INBOX from:bsnsnotify@mail.post.gov.tw subject:"存簿儲金入帳通知" -label:' + labelName;
  var threads = GmailApp.search(query);

  if (threads.length === 0) {
    console.log("目前沒有新的郵局入帳通知。");
    return;
  }

  // 嘗試在指定工作表中比對並寫入匯款資料。
  // 回傳 { found, matched, name, jumpRow, matchStatus, sheetId, colLetter }
  // found: 末五碼是否存在於該工作表；matched: 是否成功寫入（found 且有空白行）
  function tryMatch(info, data) {
    if (!info || !data.lastFive) return { found: false };

    var searchRange = info.colLetter_Last5 + ":" + info.colLetter_Last5;
    var occurrences = info.sheet.getRange(searchRange).createTextFinder(data.lastFive).matchEntireCell(true).findAll();

    if (occurrences.length === 0 && data.lastFive.startsWith("0")) {
      var numberString = parseInt(data.lastFive, 10).toString();
      occurrences = info.sheet.getRange(searchRange).createTextFinder(numberString).matchEntireCell(true).findAll();
    }

    if (occurrences.length === 0) return { found: false };

    var updatedRows = [];
    var hasFilledAmount = false;
    var name = "";
    var jumpRow = null;

    for (var k = 0; k < occurrences.length; k++) {
      var currentRow = occurrences[k].getRow();
      var cellAmount = info.sheet.getRange(currentRow, info.colIdx_Amount);
      var cellNeedAmount = info.sheet.getRange(currentRow, info.colIdx_NeedAmount);
      var cellStatus = info.sheet.getRange(currentRow, info.colIdx_Status);
      var currentAmountValue = cellAmount.getValue();

      if (currentAmountValue === "") {
        jumpRow = currentRow;

        if (!hasFilledAmount) {
          var mailAmount = parseFloat(data.amount) || 0;
          var needAmount = parseFloat(cellNeedAmount.getValue()) || 0;

          if (needAmount === mailAmount) {
            var targetStatus = "已收到錢";
            var currentStatus = String(cellStatus.getValue());
            if (currentStatus.indexOf(targetStatus) === -1) {
              cellStatus.setValue(currentStatus === "" ? targetStatus : currentStatus + "," + targetStatus);
            }
          }

          cellAmount.setValue(data.amount);
          updatedRows.push(currentRow);
          hasFilledAmount = true;

          try {
            var customerEmail = info.sheet.getRange(currentRow, info.colIdx_Email).getValue();
            var phone = info.sheet.getRange(currentRow, info.colIdx_Phone).getValue();
            name = info.sheet.getRange(currentRow, info.colIdx_Name).getValue();

            var subject = "【收款通知】喇當大叔的果園已收到您的款項 (" + data.amount + "元)";
            if (customerEmail && customerEmail.includes("@")) {
              var emailBody;
              if (needAmount === mailAmount)
                emailBody = "親愛的顧客您好：\n\n我們已經收到您的款項 " + data.amount + " 元。\n感謝您的訂購！";
              else
                emailBody = "親愛的顧客您好：\n\n我們已經收到您的款項 " + data.amount + " 元。\n訂單應付金額為" + needAmount + "請檢查匯款金額是否有誤。\n感謝您的訂購！";
              GmailApp.sendEmail(customerEmail, subject, emailBody);
              console.log(">> 對帳成功發信 (Row: " + currentRow + "): " + customerEmail);
            } else if (phone) {
              var cleanPhone = String(phone).replace(/-/g, "").trim();
              if (typeof sendTwSms === "function") {
                sendTwSms(cleanPhone, subject);
                console.log(">> 對帳成功發簡訊 (Row: " + currentRow + "): " + cleanPhone);
              }
            }
          } catch (e) {
            console.log(">> 對帳發信錯誤 (Row: " + currentRow + "): " + e.message);
          }

        } else {
          cellAmount.setValue("可能有併單匯款").setBackground("#ffcccc");
          updatedRows.push(currentRow + "(併單)");
        }
      }
    }

    var matchStatus = updatedRows.length > 0
      ? "對帳成功 (編號: " + updatedRows.join(", ") + ")"
      : "失敗：皆已入帳";

    return {
      found: true,
      name: name,
      jumpRow: jumpRow,
      matchStatus: matchStatus,
      sheetId: info.sheetId,
      colLetter: info.colLetter_Last5
    };
  }

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var message = thread.getMessages().pop();
    var body = message.getBody();

    var data = { date: "", time: "", amount: "", outAccount: "", lastFive: "", outBank: "" };

    var dateMatch = body.match(/轉入時間：[\s\S]*?(\d{3}\/\d{2}\/\d{2})/);
    if (dateMatch) {
      var parts = dateMatch[1].split('/');
      data.date = (parseInt(parts[0]) + 1911) + "/" + parts[1] + "/" + parts[2];
    }
    var timeMatch = body.match(/轉入時間：[\s\S]*?\d{3}\/\d{2}\/\d{2}[\s\S]*?(\d{2}:\d{2})/);
    if (timeMatch) data.time = timeMatch[1];

    var amountMatch = body.match(/轉入金額：[\s\S]*?>([\d,]+)元/);
    if (amountMatch) data.amount = amountMatch[1].replace(/,/g, "");

    var outAccountMatch = body.match(/轉出帳號：[\s\S]*?>([\d\*]+)</);
    if (outAccountMatch) {
      data.outAccount = outAccountMatch[1];
      data.lastFive = data.outAccount.slice(-5);
    }

    var outBankMatch = body.match(/轉出行庫：[\s\S]*?>([^<]+)</);
    if (outBankMatch) data.outBank = outBankMatch[1].replace(/&nbsp;/g, "").trim();

    // 先比對「訂單」，找不到末五碼才比對「李子訂單」
    var result = tryMatch(sheetInfo1, data);
    if (!result.found) {
      result = tryMatch(sheetInfo2, data);
    }

    var matchStatus = result.found ? result.matchStatus : "未找到訂單";
    var name = result.found ? result.name : "";
    var jumpRow = result.found ? result.jumpRow : null;

    logSheet.appendRow([
      data.date, data.time, data.amount, data.outAccount, "'" + data.lastFive,
      data.outBank, matchStatus, new Date(), name
    ]);

    if (jumpRow) {
      var lastRow = logSheet.getLastRow();
      var linkUrl = '#gid=' + result.sheetId + '&range=' + result.colLetter + jumpRow;
      logSheet.getRange(lastRow, 7).setRichTextValue(
        SpreadsheetApp.newRichTextValue().setText(matchStatus).setLinkUrl(linkUrl).build()
      );
    }

    thread.markRead().addLabel(label).moveToTrash();
  }
}
