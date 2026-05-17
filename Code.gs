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
const ORDER_SHEET    = '訂單';
const CUSTOMER_SHEET = '老客戶';
const GROCERY_SHEET = '香菇品項';

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
    const data  = sheet.getDataRange().getValues();
    // 客戶工作表欄位：電話 | 姓名 | Email | 地址
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === phone) {
        return jsonResp({
          success: true,
          customer: {
            phone:   data[i][0],
            name:    data[i][1],
            email:   data[i][2],
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
    if (!payload.phone)    return jsonResp({ success: false, message: '缺少電話' });
    if (!payload.buyerName) return jsonResp({ success: false, message: '缺少姓名' });
    if (!payload.address)  return jsonResp({ success: false, message: '缺少地址' });
    if (!payload.items || !payload.items.length) return jsonResp({ success: false, message: '無訂購品項' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── 寫入訂單工作表 ──
    const orderId = 'ORD' + new Date().getTime();
    const orderSheet = ss.getSheetByName(ORDER_SHEET);

    // 品項序列化：「特大朵 300g×2, 大朵 150g×1」
    const itemStr = payload.items
      .map(i => `${i.name} ${i.unit}×${i.qty}`)
      .join(', ');

    // 訂單工作表欄位順序
    const rowData = [
      orderId,
      payload.orderTime || new Date().toLocaleString('zh-TW'),
      payload.buyerName,
      payload.phone,
      payload.email || '',
      payload.isGift ? '是' : '否',
      payload.recipientName  || '',
      payload.recipientPhone || '',
      payload.address,
      payload.note || '',
      itemStr,
      payload.total,
      payload.transferLast5 || '',
      payload.transferBank  || '',
      '待匯款'  // 初始付款狀態
    ];

    // 先把會有前導零的欄位格式設為文字，再寫入，避免 Sheets 自動轉數字
    // clearDataValidations() 先移除可能存在的驗證規則，否則 setNumberFormat 會拋錯
    const nextRow = orderSheet.getLastRow() + 1;
    orderSheet.getRange(nextRow, 4).clearDataValidations().setNumberFormat('@');   // 電話
    orderSheet.getRange(nextRow, 8).clearDataValidations().setNumberFormat('@');   // 收件人電話
    orderSheet.getRange(nextRow, 13).clearDataValidations().setNumberFormat('@');  // 匯款末五碼
    orderSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

    // ── 扣減商品庫存 ──
    deductInventory(ss, payload.items);

    // ── 更新或新增客戶資料 ──
    updateCustomer(ss, {
      phone:   payload.phone,
      name:    payload.buyerName,
      email:   payload.email || '',
      address: payload.address
    });

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
function deductInventory(ss, items) {
  const sheet = ss.getSheetByName(GROCERY_SHEET);
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
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === customer.phone) {
      // 更新現有記錄（只更新非空欄位）
      const row = i + 1;
      sheet.getRange(row, 2).setValue(customer.name);
      if (customer.email)   sheet.getRange(row, 3).setValue(customer.email);
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
