/**
 * Finance OS 雲端同步 + 股價自動更新 — Google Apps Script
 *
 * 功能：
 *  - 儲存/讀取整份記帳資料(JSON) → data 工作表 A1
 *  - 用內建 GOOGLEFINANCE 抓股價與匯率 → ?action=quotes
 *
 * 使用步驟：
 * 1) 把下面 TOKEN 改成你自己的密碼(和 App 裡填的一模一樣)
 * 2) 部署 → 管理部署作業 → ✏️ 編輯 → 版本選「新版本」→ 部署
 *    (執行身分：我；誰可以存取：所有人)
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
  if (e.parameter.action === 'quotes') return quotes(e);
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

/** 用 GOOGLEFINANCE 抓報價；symbols 以逗號分隔，例如 TPE:2330,VOO */
function quotes(e) {
  const syms = (e.parameter.symbols || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const sh = dataSheet();
  const COL = 30; // 用第 30 欄(AD)當暫存區，不會動到 A1
  // 寫入報價公式
  syms.forEach(function (s, i) {
    sh.getRange(i + 1, COL).setFormula('=IFERROR(GOOGLEFINANCE("' + s + '","price"),"NA")');
  });
  // 匯率公式
  sh.getRange(1, COL + 1).setFormula('=IFERROR(GOOGLEFINANCE("CURRENCY:USDTWD"),"NA")');
  SpreadsheetApp.flush();
  // 讀回
  const res = {};
  syms.forEach(function (s, i) {
    res[s] = sh.getRange(i + 1, COL).getValue();
  });
  const rate = sh.getRange(1, COL + 1).getValue();
  // 清掉暫存
  sh.getRange(1, COL, Math.max(syms.length, 1), 2).clearContent();
  return out({ ok: true, quotes: res, rate: rate });
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
