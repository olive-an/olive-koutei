// 工程カレンダー 共有用（Googleスプレッドシートに貼り付けて使います）
// 予定・企業・従業員を、1件ずつ「data」シートに保存します。
var SHEET_NAME = 'data';

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(['key', 'json', 't', 'rev', 'deleted']);
    s.setFrozenRows(1);
  }
  return s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, msg: '工程カレンダーの共有は動いています' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var req = JSON.parse(e.postData.contents);
    var s = sheet_();
    var last = s.getLastRow();
    var rows = last > 1 ? s.getRange(2, 1, last - 1, 5).getValues() : [];
    var idx = {};
    var rev = Number(PropertiesService.getScriptProperties().getProperty('rev') || 0);
    rows.forEach(function (r, i) {
      idx[r[0]] = i;
      if (Number(r[3]) > rev) rev = Number(r[3]);
    });
    var since = Number(req.since) || 0;
    var rejected = {};
    var changed = false;
    (req.changes || []).forEach(function (c) {
      if (!c || typeof c.k !== 'string') return;
      var i = idx[c.k];
      if (i !== undefined && Number(rows[i][2]) > Number(c.t)) { rejected[c.k] = 1; return; }
      rev++;
      var row = [c.k, c.d ? '' : String(c.j || ''), Number(c.t) || Date.now(), rev, c.d ? 1 : 0];
      if (i !== undefined) rows[i] = row;
      else { rows.push(row); idx[c.k] = rows.length - 1; }
      changed = true;
    });
    if (changed) {
      s.getRange(2, 1, rows.length, 5).setValues(rows);
      PropertiesService.getScriptProperties().setProperty('rev', String(rev));
    }
    var recs = [];
    rows.forEach(function (r) {
      if (Number(r[3]) > since || rejected[r[0]]) {
        recs.push({ k: r[0], j: String(r[1]), t: Number(r[2]), r: Number(r[3]), d: Number(r[4]) ? 1 : 0 });
      }
    });
    return out_({ ok: true, rev: rev, recs: recs });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
