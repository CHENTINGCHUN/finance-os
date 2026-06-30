/**
 * Finance OS 雲端同步 — Google Apps Script
 * 把整份記帳資料(JSON)存到這份試算表的 data 工作表 A1 儲存格。
 *
 * 使用步驟見 README，重點：
 * 1) 把下面 TOKEN 改成你自己的密碼(英數字皆可)
 * 2) 部署 → 新增部署作業 → 類型「網頁應用程式」
 *    執行身分：我
 *    誰可以存取：任何人
 * 3) 複製產生的「網頁應用程式網址」(以 /exec 結尾)
 * 4) 在 Finance OS 的 ⚙ 設定裡貼上網址 + 一樣的 TOKEN
 */

const TOKEN = 'change-me-123';  // ← 改成你自己的密碼，和 App 裡填的要一模一樣

function dataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('data');
  if (!sh) sh = ss.insertSheet('data');
  return sh;
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== TOKEN) return out({ error: 'unauthorized' });
  const v = dataSheet().getRange('A1').getValue();
  return out({ ok: true, data: v });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return out({ error: 'bad json' }); }
  if (body.token !== TOKEN) return out({ error: 'unauthorized' });
  dataSheet().getRange('A1').setValue(body.data || '');
  return out({ ok: true });
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
