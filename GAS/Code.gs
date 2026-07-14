// ════════════════════════════════════════════════════
// 喇當大叔的果園 — 訂單 GAS 腳本
// ════════════════════════════════════════════════════
//
// 試算表結構（請先建好以下工作表）：
//   工作表1「訂單」   ← 每筆訂單寫一列
//   工作表2「客戶」   ← 老客戶查詢用（電話 / 姓名 / Email / 地址）
//
// 部署方式：
//   擴充功能 → Apps Script → 部署 → 新增部署
//   類型：Web 應用程式
//   執行身分：我（你的帳號）
//   存取權限：任何人
// ════════════════════════════════════════════════════

const SPREADSHEET_ID = '13s8MhC6LmZSA0Brbal7bigI_n4QaJuqWMiSg6BfL1kk';   // ← 換成你的試算表 ID
const ORDER_SHEET = '訂單';
const PLUM_ORDER_SHEET = '李子訂單';
const PEACH_ORDER_SHEET = '水蜜桃訂單';
const CUSTOMER_SHEET = '老客戶';
const GROCERY_SHEET = '香菇品項';
const PLUM_GROCERY_SHEET = '李子品項';
const PEACH_GROCERY_SHEET = '水蜜桃品項';

// ────────────────────────────────────────
// doGet：老客戶查詢
// 呼叫範例：?action=getCustomer&phone=0912345678
// ────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getCustomer') {
    const phone = (e.parameter.phone || '').trim();
    if (!phone) return jsonResp({ success: false, message: '未提供電話' });

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CUSTOMER_SHEET);
    const data = sheet.getDataRange().getValues();
    // 客戶工作表欄位：電話 | 姓名 | Email | 地址
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === phone) {
        return jsonResp({
          success: true,
          customer: {
            phone: data[i][0],
            name: data[i][1],
            email: data[i][2],
            address: data[i][3]
          }
        });
      }
    }
    return jsonResp({ success: false, message: '查無此電話' });
  }

  return jsonResp({ success: false, message: 'unknown action' });
}

// ────────────────────────────────────────
// doPost：接收訂單，寫入試算表
// ────────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);  // 最多等 15 秒，避免同時寫入衝突

  try {
    // 前端用 text/plain 傳 JSON，要自己 parse
    const payload = JSON.parse(e.postData.contents);

    // ── 後端驗證 ──
    if (!payload.phone) return jsonResp({ success: false, message: '缺少電話' });
    if (!payload.buyerName) return jsonResp({ success: false, message: '缺少姓名' });
    if (!payload.address && !payload.isPickup) return jsonResp({ success: false, message: '缺少地址' });
    if (!payload.items || !payload.items.length) return jsonResp({ success: false, message: '無訂購品項' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── 寫入訂單工作表（依 orderType 分流）──
    const isPlum = payload.orderType === 'plum';
    const isPeach = payload.orderType === 'peach';
    const orderId = 'ORD' + new Date().getTime();
    const orderSheet = ss.getSheetByName(isPeach ? PEACH_ORDER_SHEET : (isPlum ? PLUM_ORDER_SHEET : ORDER_SHEET));

    // 品項序列化：「特大朵 300g×2, 大朵 150g×1」
    const itemStr = payload.items
      .map(i => `${i.name} ${i.unit}×${i.qty}`)
      .join(', ');

    // 收件 / 付款方式
    const isPickup = !!payload.isPickup;
    const isCod = !isPickup && !!payload.isCod;
    const payStatus = isPickup ? '面交付款' : (isCod ? '貨到付款' : '待匯款');
    const deliveryTime = isPickup ? '' : (payload.deliveryTime ? Number(payload.deliveryTime) : '');

    // 訂單工作表欄位順序
    // A~O 為原有欄位；P~S 為新增欄位（收件方式 / 配送時段 / 付款方式 / 物流處理費）
    const rowData = [
      orderId,
      payload.orderTime || new Date().toLocaleString('zh-TW'),
      payload.buyerName,
      payload.phone,
      payload.email || '',
      payload.isGift ? '是' : '否',
      payload.recipientName || '',
      payload.recipientPhone || '',
      payload.address || '',
      payload.note || '',
      itemStr,
      payload.total,
      payload.transferLast5 || '',
      payload.transferBank || '',
      payStatus,                       // O 付款狀態
      isPickup ? '埔里自取' : '宅配',   // P 收件方式
      deliveryTime,                    // Q 配送時段（1/2/4）
      isCod ? '貨到付款' : '匯款',      // R 付款方式
      payload.codFee || payload.shippingFee || 0   // S 物流處理費／運費（水蜜桃為宅配運費）
    ];

    // 以 A 欄最後一筆有值的列號 +1 決定寫入位置，
    // 避免右側處理表的資料影響 getLastRow() 的結果
    const colAValues = orderSheet.getRange('A:A').getValues();
    let nextRow = 1;
    for (let i = colAValues.length - 1; i >= 0; i--) {
      if (colAValues[i][0] !== '') { nextRow = i + 2; break; }
    }
    orderSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    // 此工作表的「輸入欄」無法用 setNumberFormat 設格式；
    // 改寫 ="value" 公式讓 Sheets 以文字儲存，保留電話前導零
    const asText = v => `="${String(v || '').replace(/"/g, '""')}"`;
    orderSheet.getRange(nextRow, 4).setFormula(asText(payload.phone));
    orderSheet.getRange(nextRow, 8).setFormula(asText(payload.recipientPhone || ''));
    orderSheet.getRange(nextRow, 13).setFormula(asText(payload.transferLast5 || ''));

    // ── 扣減商品庫存 ──
    deductInventory(ss, payload.items, isPeach ? PEACH_GROCERY_SHEET : (isPlum ? PLUM_GROCERY_SHEET : GROCERY_SHEET));

    // ── 更新或新增客戶資料 ──
    updateCustomer(ss, {
      phone: payload.phone,
      name: payload.buyerName,
      email: payload.email || '',
      address: payload.address
    });

    //發通知信件
    // 1. 檢查 Email 格式是否正確
    var emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    if (emailRegex.test(payload.email)) {
      // 格式正確，發送 HTML 信件（依 orderType 選擇模板：plum / peach / 預設香菇）
      sendOrderEmail_HTML(payload.email, orderId, payload.orderType || '');
    }
    else {
      // 2. 格式不正確，改發簡訊
      var cleanPhone = String(payload.phone).replace(/-/g, "").replace(/\s/g, "").trim();
      var message = "【訂單成立】喇當大叔的果園已經收到您的訂單。感謝您的支持！";
      sendTwSms(cleanPhone, message);
    }

    return jsonResp({ success: true, orderId: orderId });

  } catch (err) {
    console.error(err);
    return jsonResp({ success: false, message: err.toString() });

  } finally {
    lock.releaseLock();
  }
}

// ────────────────────────────────────────
// 扣減庫存：依 id 找到 GROCERY_SHEET 對應列，E欄減去訂購數量（最低為 0）
// ────────────────────────────────────────
function deductInventory(ss, items, sheetName) {
  const sheet = ss.getSheetByName(sheetName || GROCERY_SHEET);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  // 建立 id → 列號（1-based）的對照表
  const idToRow = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    if (id) idToRow[id] = i + 1;
  }

  for (const item of items) {
    const rowNum = idToRow[String(item.id).trim()];
    if (!rowNum) continue;

    const current = Number(data[rowNum - 1][4]); // E欄（index 4）
    if (isNaN(current)) continue;

    sheet.getRange(rowNum, 5).setValue(Math.max(0, current - item.qty));
  }
}

// ────────────────────────────────────────
// 客戶資料：有就更新，沒有就新增
// ────────────────────────────────────────
function updateCustomer(ss, customer) {
  const sheet = ss.getSheetByName(CUSTOMER_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === customer.phone) {
      // 更新現有記錄（只更新非空欄位）
      const row = i + 1;
      sheet.getRange(row, 2).setValue(customer.name);
      if (customer.email) sheet.getRange(row, 3).setValue(customer.email);
      if (customer.address) sheet.getRange(row, 4).setValue(customer.address);
      return;
    }
  }

  // 新增客戶
  sheet.appendRow([customer.phone, customer.name, customer.email, customer.address]);
}

// ────────────────────────────────────────
// 共用：回傳 JSON 回應
// ────────────────────────────────────────
function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
