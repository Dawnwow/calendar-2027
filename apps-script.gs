// 2027 年曆預購登記 — Google Apps Script 後端
// 貼到 Google Apps Script 編輯器，部署為 Web App
// 注意：Google Sheets「庫存」Tab 的 B 欄品項名稱，筆記本欄位填「筆記本」（不含「品相」）

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'stock';
  if (action === 'stock') {
    return jsonResponse(getStock());
  }
  return jsonResponse({ ok: false, msg: 'unknown action' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const payload = JSON.parse(e.postData.contents);
    const name  = payload.name;
    const phone = payload.phone;
    const note  = payload.note || '';
    const items = payload.items; // [{ pid, v, qty }]

    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const stockSheet = ss.getSheetByName('庫存');
    const orderSheet = ss.getSheetByName('訂單');
    const stockData  = stockSheet.getDataRange().getValues();

    // 建立庫存 index（key = pid_variant）
    const rowMap = {};
    for (let i = 1; i < stockData.length; i++) {
      const key = stockData[i][0] + '_' + stockData[i][2];
      rowMap[key] = {
        row:   i + 1,
        pname: stockData[i][1],
        stock: Number(stockData[i][3])
      };
    }

    // 驗證庫存（送出前再確認）
    for (const item of items) {
      const key  = item.pid + '_' + item.v;
      const info = rowMap[key];
      if (!info) {
        return jsonResponse({ ok: false, msg: '找不到品項：' + key });
      }
      if (info.stock < item.qty) {
        return jsonResponse({
          ok: false,
          msg: info.pname + ' 款式 ' + item.v + ' 庫存不足，剩餘 ' + info.stock + ' 本'
        });
      }
    }

    // 扣庫存 + 寫訂單
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
    for (const item of items) {
      const key  = item.pid + '_' + item.v;
      const info = rowMap[key];
      stockSheet.getRange(info.row, 4).setValue(info.stock - item.qty);
      orderSheet.appendRow([now, name, phone, info.pname, item.v, item.qty, note]);
    }

    return jsonResponse({ ok: true });

  } catch (err) {
    return jsonResponse({ ok: false, msg: err.message });
  } finally {
    lock.releaseLock();
  }
}

function getStock() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('庫存');
  const data  = sheet.getDataRange().getValues();
  const inv   = {};
  for (let i = 1; i < data.length; i++) {
    const pid = data[i][0], v = data[i][2], qty = Number(data[i][3]);
    if (!inv[pid]) inv[pid] = {};
    inv[pid][v] = qty;
  }
  return inv;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
