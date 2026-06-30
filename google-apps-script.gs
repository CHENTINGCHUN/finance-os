/**
 * Finance OS 雲端同步 + 股價自動更新 — Google Apps Script
 *
 * 功能：
 *  - 儲存/讀取整份記帳資料(JSON) → data 工作表 A1
 *  - 用 Yahoo 財經抓股價與匯率 → ?action=quotes
 *    台股自動先試上市(.TW)再試上櫃(.TWO)，覆蓋率高，連威剛(3260,上櫃)都抓得到
 *
 * 改完記得：部署 → 管理部署作業 → ✏️ 編輯 → 版本選「新版本」→ 部署
 *           (執行身分：我；誰可以存取：所有人)
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

/** 抓報價：symbols 為原始代號逗號分隔，例如 2330,3260,VOO */
function quotes(e) {
  const syms = (e.parameter.symbols || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const res = {};
  syms.forEach(function (sym) { res[sym] = priceOf(sym); });
  const rate = yahoo('USDTWD=X');
  return out({ ok: true, quotes: res, rate: (rate || 'NA') });
}

/** 判斷台股(數字代號)→先上市再上櫃；其餘當美股 */
function priceOf(sym) {
  if (/^\d+[A-Za-z]?$/.test(sym)) {
    return yahoo(sym + '.TW') || yahoo(sym + '.TWO') || 'NA';
  }
  return yahoo(sym) || 'NA';
}

/** 呼叫 Yahoo 財經 chart API 取得最新價 */
function yahoo(ysym) {
  try {
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(ysym) + '?interval=1d&range=1d';
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
    });
    if (resp.getResponseCode() !== 200) return null;
    const j = JSON.parse(resp.getContentText());
    const meta = j && j.chart && j.chart.result && j.chart.result[0] && j.chart.result[0].meta;
    const p = meta && meta.regularMarketPrice;
    return (typeof p === 'number' && p > 0) ? p : null;
  } catch (err) { return null; }
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
