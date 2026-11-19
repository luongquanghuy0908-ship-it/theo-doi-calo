var SS_ID = '1O6t_8Y_nl7M3HPwq8WPM-qeKggpfgO-RPlDvimwZtxs';
var H = ['Ngày', 'Calo', 'Đạm (g)', 'Số bữa', 'Đạt', 'Cân (kg)', 'Dữ liệu', 'u'];
var MIN_V = 2; // trang bản cũ (chưa gửi v) không được ghi nữa
function ss_() { return SpreadsheetApp.openById(SS_ID); }
function sheet_(name, header) {
  var ss = ss_(), s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); if (header) { s.getRange('A:A').setNumberFormat('@'); s.appendRow(header); s.setFrozenRows(1); } }
  return s;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function auth_(k) {
  if (!k) return false;
  var p = PropertiesService.getScriptProperties(), cur = p.getProperty('KEY');
  if (!cur) { p.setProperty('KEY', String(k)); return true; }
  return cur === String(k);
}
function norm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var d = new Date(s);
  return isNaN(d) ? s : Utilities.formatDate(d, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}
// Đọc tất cả dòng, mỗi ngày chỉ giữ 1 dòng có u lớn nhất
function rows_() {
  var s = sheet_('Theo dõi', H), v = s.getDataRange().getDisplayValues(), m = {};
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    var d = norm_(v[i][0]), r = v[i].slice(0, 8); r[0] = d;
    var u = Number(r[7]) || 0;
    if (!m[d] || u >= (Number(m[d][7]) || 0)) m[d] = r;
  }
  return m;
}
function write_(m) {
  var s = sheet_('Theo dõi', H), keys = Object.keys(m).sort(), data = keys.map(function (k) { return m[k]; });
  s.clearContents();
  s.getRange('A:A').setNumberFormat('@');
  s.getRange(1, 1, 1, 8).setValues([H]);
  if (data.length) s.getRange(2, 1, data.length, 8).setValues(data);
}
function days_(m) {
  var days = {};
  for (var d in m) { if (m[d][6]) { try { days[d] = JSON.parse(m[d][6]); } catch (x) {} } }
  return days;
}
function readAll_() {
  var c = sheet_('Cài đặt'), cv = c.getRange('A1:B1').getValues()[0];
  return { ok: true, v: MIN_V, cfg: cv[0] ? JSON.parse(cv[0]) : null, cfgU: Number(cv[1]) || 0, days: days_(rows_()) };
}
function doGet(e) {
  if (!auth_(e.parameter.key)) return out_({ error: 'key' });
  return out_(readAll_());
}
function doPost(e) {
  var b = JSON.parse(e.postData.contents);
  if (!auth_(b.key)) return out_({ error: 'key' });
  if (!b.v || b.v < MIN_V) return out_({ error: 'old' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var m = rows_(), rows = b.rows || [];
    for (var j = 0; j < rows.length; j++) {
      var r = rows[j].slice(0, 8), d = norm_(r[0]); r[0] = d;
      if (!m[d] || (Number(r[7]) || 0) >= (Number(m[d][7]) || 0)) m[d] = r;
    }
    write_(m);
    if (b.cfg) { var c = sheet_('Cài đặt'), old = Number(c.getRange('B1').getValue()) || 0; if ((b.cfgU || 0) >= old) c.getRange('A1:B1').setValues([[JSON.stringify(b.cfg), b.cfgU || 0]]); }
    return out_({ ok: true, days: days_(m) });
  } finally { lock.releaseLock(); }
}
function donDongTrung() { write_(rows_()); }
