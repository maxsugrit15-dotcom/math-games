/**
 * Tutor Max — อันดับออนไลน์สำหรับเกม "มี…รูปอะไรบ้าง"
 * วางโค้ดนี้ใน Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app
 * (Execute as: Me · Who has access: Anyone)
 */

var SHEET = 'scores';
var ADMIN_PIN = '2468';        // <-- เปลี่ยนเลขนี้เป็นของครูเอง ใช้ตอนสั่งล้างข้อมูล

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.cb || 'cb';
  var out;
  try { out = handle(p); }
  catch (err) { out = { ok: false, err: 'server', detail: String(err) }; }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(out) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['seat', 'name', 'av', 'score', 'ms', 'errs', 'stars', 'updated']);
  }
  return sh;
}

function readAll_() {
  var v = sheet_().getDataRange().getValues();
  v.shift();
  return v.filter(function (r) { return String(r[0]).length > 0; })
          .map(function (r) {
            return { seat: String(r[0]), name: String(r[1]), av: String(r[2]),
                     score: Number(r[3]) || 0, ms: Number(r[4]) || 0,
                     errs: Number(r[5]) || 0, stars: Number(r[6]) || 0 };
          });
}

function rowOfSeat_(seat) {
  var v = sheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(seat)) return i + 1;
  return 0;
}

function handle(p) {
  var a = p.action;

  if (a === 'list') return { ok: true, players: readAll_() };

  if (a === 'claim') {
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var seat = String(p.seat || ''), name = String(p.name || '').trim(), av = String(p.av || '');
      if (!seat || !name || !av) return { ok: false, err: 'bad_input' };
      var all = readAll_();
      var others = all.filter(function (x) { return x.seat !== seat; });
      var lower = name.toLowerCase();
      if (others.some(function (x) { return x.name.toLowerCase() === lower; }))
        return { ok: false, err: 'name', players: all };
      if (others.some(function (x) { return x.av === av; }))
        return { ok: false, err: 'av', players: all };
      var mine = all.filter(function (x) { return x.seat === seat; })[0];
      if (mine && mine.name.toLowerCase() !== lower)
        return { ok: false, err: 'seat', by: mine.name, players: all };
      if (!mine) sheet_().appendRow([seat, name, av, '', '', '', '', new Date()]);
      else {
        var row = rowOfSeat_(seat);
        sheet_().getRange(row, 2, 1, 2).setValues([[name, av]]);
        sheet_().getRange(row, 8).setValue(new Date());
      }
      return { ok: true, players: readAll_() };
    } finally { lock.releaseLock(); }
  }

  if (a === 'score') {
    var lock2 = LockService.getScriptLock();
    lock2.waitLock(15000);
    try {
      var seat2 = String(p.seat || ''), nm = String(p.name || '').trim();
      var sc = Number(p.score) || 0, ms = Number(p.ms) || 0;
      var er = Number(p.errs) || 0, st = Number(p.stars) || 0;
      if (!seat2 || !nm) return { ok: false, err: 'bad_input' };
      var row2 = rowOfSeat_(seat2);
      if (!row2) return { ok: false, err: 'not_claimed' };
      var sh = sheet_();
      var cur = sh.getRange(row2, 1, 1, 8).getValues()[0];
      if (String(cur[1]).toLowerCase() !== nm.toLowerCase())
        return { ok: false, err: 'wrong_player', players: readAll_() };
      var curScore = Number(cur[3]) || 0, curMs = Number(cur[4]) || 0;
      var better = !curScore || sc > curScore || (sc === curScore && ms < curMs);
      if (better) {
        sh.getRange(row2, 4, 1, 5).setValues([[sc, ms, er, st, new Date()]]);
      }
      return { ok: true, saved: better, players: readAll_() };
    } finally { lock2.releaseLock(); }
  }

  if (a === 'reset') {
    if (String(p.pin) !== String(ADMIN_PIN)) return { ok: false, err: 'pin' };
    var sh2 = sheet_();
    if (sh2.getLastRow() > 1) sh2.deleteRows(2, sh2.getLastRow() - 1);
    return { ok: true, players: [] };
  }

  return { ok: false, err: 'unknown_action' };
}
