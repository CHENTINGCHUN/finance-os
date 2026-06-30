/**
 * Finance OS 雲端同步 + 股價自動更新 — Google Apps Script
 *
 * 報價多來源自動接力（解決 Yahoo 擋 Apps Script 機房的問題）：
 *   1) GOOGLEFINANCE（Google 自家，一定連得到）
 *   2) 證交所/櫃買 MIS 官方 API（涵蓋上市 tse_ 與上櫃 otc_）
 *   3) Yahoo 財經（備援）
 * 回傳會附 sources 欄位，顯示每檔是用哪個來源抓到的，方便診斷。
 *
 * 改完記得：部署 → 管理部署作業 → ✏️ 編輯 → 版本選「新版本」→ 部署
 *           (執行身分：我；誰可以存取：所有人)；若問外部連線權限，按允許。
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

function isTW(sym) { return /^\d+[A-Za-z]?$/.test(sym); }

/** 抓報價：symbols 為原始代號逗號分隔，例如 2330,3260,VOO */
function quotes(e) {
  const syms = (e.parameter.symbols || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const gf = gfBatch(syms);          // 1) 先批次 GOOGLEFINANCE
  const res = {}, src = {};
  syms.forEach(function (sym) {
    if (typeof gf[sym] === 'number') { res[sym] = gf[sym]; src[sym] = 'GF'; return; }
    if (isTW(sym)) {                  // 2) 台股 → 證交所/櫃買官方
      var p = twse(sym);
      if (p) { res[sym] = p; src[sym] = 'TWSE'; return; }
      var y = yahoo(sym + '.TW') || yahoo(sym + '.TWO');  // 3) Yahoo 備援
      if (y) { res[sym] = y; src[sym] = 'YH'; return; }
    } else {
      var u = yahoo(sym);
      if (u) { res[sym] = u; src[sym] = 'YH'; return; }
    }
    res[sym] = 'NA'; src[sym] = '-';
  });
  const rate = gfRate() || yahoo('USDTWD=X') || 'NA';
  return out({ ok: true, quotes: res, sources: src, rate: rate });
}

/** 批次 GOOGLEFINANCE（暫存在第 30 欄，不動到 A1） */
function gfBatch(syms) {
  const sh = dataSheet(), COL = 30;
  syms.forEach(function (sym, i) {
    var t = isTW(sym) ? ('TPE:' + sym) : sym;
    sh.getRange(i + 1, COL).setFormula('=IFERROR(GOOGLEFINANCE("' + t + '","price"),"NA")');
  });
  SpreadsheetApp.flush();
  const res = {};
  syms.forEach(function (sym, i) {
    var v = sh.getRange(i + 1, COL).getValue();
    res[sym] = (typeof v === 'number' && v > 0) ? v : null;
  });
  sh.getRange(1, COL, Math.max(syms.length, 1), 1).clearContent();
  return res;
}

/** GOOGLEFINANCE 匯率 */
function gfRate() {
  const c = dataSheet().getRange(1, 32);
  c.setFormula('=IFERROR(GOOGLEFINANCE("CURRENCY:USDTWD"),"NA")');
  SpreadsheetApp.flush();
  const v = c.getValue(); c.clearContent();
  return (typeof v === 'number' && v > 0) ? v : null;
}

/** 證交所/櫃買 MIS 官方 API：先試上市 tse_，再試上櫃 otc_ */
function twse(code) {
  const chs = ['tse_' + code + '.tw', 'otc_' + code + '.tw'];
  for (var i = 0; i < chs.length; i++) {
    try {
      var url = 'https://mis.twse.com.tw/stock/api/getStockInfo.jsp?json=1&delay=0&ex_ch=' + chs[i];
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (resp.getResponseCode() !== 200) continue;
      var j = JSON.parse(resp.getContentText());
      if (j && j.msgArray && j.msgArray.length) {
        var d = j.msgArray[0];
        var p = parseFloat(d.z);                 // 最新成交價
        if (!(p > 0)) p = parseFloat(d.y);       // 昨收備用
        if (!(p > 0)) {                           // 用買賣價估
          var a = parseFloat((d.a || '').split('_')[0]);
          var b = parseFloat((d.b || '').split('_')[0]);
          if (a > 0 && b > 0) p = (a + b) / 2; else p = a > 0 ? a : b;
        }
        if (p > 0) return p;
      }
    } catch (err) { }
  }
  return null;
}

/** Yahoo 財經 chart API */
function yahoo(ysym) {
  try {
    var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(ysym) + '?interval=1d&range=1d';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' } });
    if (resp.getResponseCode() !== 200) return null;
    var j = JSON.parse(resp.getContentText());
    var meta = j && j.chart && j.chart.result && j.chart.result[0] && j.chart.result[0].meta;
    var p = meta && meta.regularMarketPrice;
    return (typeof p === 'number' && p > 0) ? p : null;
  } catch (err) { return null; }
}

function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
