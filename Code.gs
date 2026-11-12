/**
 * HÀNH TRÌNH CỦA QUANG ĐẠT – Apps Script nhận kết quả & xem lại bài làm
 * ---------------------------------------------------------------------
 * CÁCH CÀI (làm 1 lần):
 * 1. Tạo một Google Sheet mới (vd: "Kết quả ôn thi – Quang Đạt").
 * 2. Trong Sheet: Tiện ích mở rộng (Extensions) → Apps Script.
 *    Xóa hết code mẫu, dán toàn bộ file này vào, bấm Lưu.
 * 3. Đổi VIEW_PIN bên dưới thành mã riêng của thầy (mã để mở trang xem kết quả).
 * 4. Chọn hàm caiDat ở thanh trên → bấm Chạy → cấp quyền (Allow).
 * 5. Triển khai (Deploy) → Triển khai mới (New deployment) → loại "Ứng dụng web" (Web app)
 *      - Thực thi dưới tên (Execute as): Tôi (Me)
 *      - Người có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 *    → Triển khai → copy "URL ứng dụng web" (đuôi /exec).
 * 6. Dán URL đó vào dòng SCRIPT_URL ở đầu file toan.html và tieng-anh.html.
 * 7. Xem kết quả: mở  URL_ứng_dụng_web?pin=MÃ_CỦA_THẦY
 *
 * Lưu ý: nếu sau này sửa code này, phải Deploy → Quản lý triển khai → sửa (bút chì)
 * → Phiên bản: "Phiên bản mới" thì code mới có hiệu lực, link /exec giữ nguyên.
 */

const VIEW_PIN = '2026';            // mã để xem kết quả – thầy nên đổi
const SHEET_ID = '';                // (không bắt buộc) ID Google Sheet muốn ghi vào. Để trống: tự dùng Sheet gắn với script, hoặc tự tạo Sheet mới
const TZ = 'Asia/Ho_Chi_Minh';
const SH_ATTEMPT = 'Lượt làm';
const SH_DETAIL = 'Chi tiết từng câu';
const SH_JSON = 'Dữ liệu (không sửa)';

const H_ATTEMPT = ['Thời gian nộp', 'Mã lượt', 'Học sinh', 'Môn', 'Thử thách số', 'Tên thử thách', 'Nội dung', 'Trạng thái',
  'Đúng ngay lần đầu', 'Tổng số câu', 'Tỉ lệ %', 'Số sao', 'Tiền vàng', 'Thời gian làm', 'Số câu nhờ gợi ý', 'Số lần hết tim', 'Xem lại bài'];
const H_DETAIL = ['Mã lượt', 'Thời gian nộp', 'Môn', 'Thử thách', 'Câu', 'Đề bài', 'Đạt trả lời (các lần)', 'Đáp án đúng', 'Kết quả', 'Gợi ý đã mở'];

/* ============ CÀI ĐẶT BAN ĐẦU ============ */
function caiDat() {
  const ss = ss_();
  sheet_(SH_ATTEMPT, H_ATTEMPT);
  sheet_(SH_DETAIL, H_DETAIL);
  const j = sheet_(SH_JSON, ['Mã lượt', 'JSON']);
  j.hideSheet();
  const s = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('Trang tính1');
  if (s && s.getLastRow() === 0 && ss_().getSheets().length > 3) ss_().deleteSheet(s);
  const msg = 'Đã cài đặt xong. Kết quả sẽ ghi vào Sheet: ' + ss.getUrl();
  Logger.log(msg);
  return msg;
}

/** Chạy hàm này trong trình soạn thảo để thử gửi 1 kết quả mẫu và xem có lỗi gì không */
function guiThu() {
  const d = { id: 'TEST' + Date.now(), student: 'Thử nghiệm', subject: 'Thử nghiệm', subjectKey: 'test', level: 0, levelTitle: 'Gửi thử', levelSub: '', status: 'Thử',
    startedAt: new Date().toISOString(), finishedAt: new Date().toISOString(), seconds: 1, correctFirst: 1, total: 1, percent: 100, stars: 3, coins: 10, hintsUsed: 0, revives: 0,
    details: [{ n: 1, q: 'Câu thử', type: 'in', answers: ['1'], correct: '1', first: true, final: true, hints: 0, hintTotal: 0 }] };
  const r = doPost({ postData: { contents: JSON.stringify(d) } }).getContent();
  Logger.log('Kết quả: ' + r + ' | Sheet: ' + ss_().getUrl());
  return r;
}
function ss_() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { }
  if (ss) return ss;
  // script đứng riêng (không gắn với Sheet): tự tạo 1 Sheet và nhớ ID của nó
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('Kết quả ôn thi – Quang Đạt');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function sheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2b2770').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============ NHẬN KẾT QUẢ TỪ GAME ============ */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    if (d && d.action === 'save') return saveProgress_(d);
    if (!d || !d.id || !Array.isArray(d.details)) return out_('bad request');
    const shA = sheet_(SH_ATTEMPT, H_ATTEMPT), shD = sheet_(SH_DETAIL, H_DETAIL), shJ = sheet_(SH_JSON, ['Mã lượt', 'JSON']);
    // chống ghi trùng khi game gửi lại
    if (shJ.getLastRow() > 1) {
      const ids = shJ.getRange(2, 1, shJ.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === d.id) return out_('duplicate');
    }
    const now = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
    const link = ScriptApp.getService().getUrl() + '?pin=' + encodeURIComponent(VIEW_PIN) + '&id=' + encodeURIComponent(d.id);
    shA.appendRow([now, d.id, str_(d.student), str_(d.subject).split('·')[0].trim(), d.level, str_(d.levelTitle), str_(d.levelSub), str_(d.status),
      num_(d.correctFirst), num_(d.total), num_(d.percent), num_(d.stars), num_(d.coins), dur_(d.seconds), num_(d.hintsUsed), num_(d.revives), '']);
    const r = shA.getLastRow();
    shA.getRange(r, 17).setFormula('=HYPERLINK("' + link + '","Xem bài")');
    const color = d.percent >= 90 ? '#d9f7e3' : d.percent >= 70 ? '#fff4cc' : '#ffe1e5';
    shA.getRange(r, 11).setBackground(color);

    const rows = d.details.map(function (x) {
      return [d.id, now, str_(d.subject).split('·')[0].trim(), str_(d.levelTitle), x.n, plain_(x.q).slice(0, 2000),
        (x.answers || []).map(function (a, i) { return 'Lần ' + (i + 1) + ': ' + plain_(a); }).join('\n').slice(0, 5000),
        plain_(x.correct), x.type === 'write' ? 'Bài viết (thầy chấm)' : x.first ? 'Đúng ngay' : x.final ? 'Đúng lần 2' : 'Sai',
        (x.hints || 0) + '/' + (x.hintTotal || 0)];
    });
    if (rows.length) shD.getRange(shD.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    // lưu JSON đầy đủ để dựng lại giao diện bài làm (chia nhỏ nếu dài)
    const js = JSON.stringify(d), parts = [d.id];
    for (let i = 0; i < js.length; i += 45000) parts.push(js.slice(i, i + 45000));
    shJ.getRange(shJ.getLastRow() + 1, 1, 1, parts.length).setValues([parts]);
    return out_('ok');
  } catch (err) {
    try { PropertiesService.getScriptProperties().setProperty('LAST_ERROR', new Date().toISOString() + ' – ' + err); } catch (x) { }
    return out_('error: ' + err);
  } finally {
    try { lock.releaseLock(); } catch (x) { }
  }
}
/* ============ LƯU / TẢI TIẾN TRÌNH CHƠI (xu, cấp, đồ đã mua...) ============ */
const SH_SAVE = 'Tiến trình chơi';
const H_SAVE = ['Người chơi', 'Học sinh', 'Cập nhật', 'Lượt lưu', 'Cấp', 'Xu', 'Dữ liệu (không sửa)'];
function saveProgress_(d) {
  const sh = sheet_(SH_SAVE, H_SAVE), key = str_(d.player);
  if (!key || !d.save) return out_('bad request');
  let row = 0;
  if (sh.getLastRow() > 1) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < v.length; i++) if (v[i][0] === key) { if (Number(v[i][3]) > Number(d.rev)) return out_('older'); row = i + 2; break; }
  }
  const js = JSON.stringify(d.save), vals = [key, str_(d.student), Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'), num_(d.rev), num_(d.level), num_(d.coins)];
  for (let i = 0; i < js.length; i += 45000) vals.push(js.slice(i, i + 45000));
  if (!row) row = sh.getLastRow() + 1;
  else sh.getRange(row, 1, 1, Math.max(sh.getLastColumn(), vals.length)).clearContent();
  sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  return out_('saved');
}
function loadProgress_(player) {
  let save = null;
  try {
    const sh = ss_().getSheetByName(SH_SAVE);
    if (sh && sh.getLastRow() > 1) {
      const v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      for (let i = 0; i < v.length; i++) if (v[i][0] === String(player)) { save = JSON.parse(v[i].slice(6).join('')); break; }
    }
  } catch (e) { }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, save: save })).setMimeType(ContentService.MimeType.JSON);
}
function out_(t) { return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT); }
function str_(v) { return v == null ? '' : String(v).slice(0, 500); }
function num_(v) { const n = Number(v); return isFinite(n) ? n : ''; }
function dur_(s) { s = Number(s) || 0; return Math.floor(s / 60) + ' phút ' + (s % 60) + ' giây'; }
function plain_(h) {
  return String(h == null ? '' : h).replace(/<sup>(.*?)<\/sup>/g, '^$1').replace(/<br\s*\/?>/g, '\n').replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&').trim();
}

/* ============ TRANG XEM KẾT QUẢ ============ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'load') return loadProgress_(p.player || '');
  const base = ScriptApp.getService().getUrl();
  let body;
  if (String(p.pin || '') !== VIEW_PIN) body = pinPage_(base, p.pin);
  else if (p.id) body = reviewPage_(base, p.id);
  else body = dashPage_(base, p.mon || '');
  return HtmlService.createHtmlOutput(page_(body)).setTitle('Kết quả của Quang Đạt')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function page_(body) {
  return '<!doctype html><html><head><base target="_top"><meta charset="utf-8"><style>' + CSS_ + '</style></head><body><div class="wrap">' + body + '</div></body></html>';
}
const CSS_ = [
  'body{margin:0;font-family:Nunito,Segoe UI,Roboto,Arial,sans-serif;background:#f3f1fb;color:#1b1740}',
  '.wrap{max-width:980px;margin:0 auto;padding:16px}',
  'h1{font-size:26px;margin:4px 0 2px}h2{font-size:20px;margin:22px 0 8px}',
  '.muted{color:#6b678f;font-size:14px}',
  'a{color:#5b2fd6;font-weight:700}',
  '.tabs a{display:inline-block;padding:7px 14px;border-radius:20px;background:#fff;margin:8px 6px 0 0;text-decoration:none;border:2px solid #e2ddf7}',
  '.tabs a.on{background:#2b2770;color:#fff;border-color:#2b2770}',
  '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px}',
  '.card{background:#fff;border-radius:14px;padding:12px;box-shadow:0 3px 0 #e2ddf7}',
  '.card b{font-size:15px}.big{font-size:28px;font-weight:900}',
  '.bar{height:8px;border-radius:5px;background:#eee;margin-top:6px;overflow:hidden}.bar i{display:block;height:100%;background:linear-gradient(90deg,#ffcf33,#ff5d73)}',
  'table{width:100%;border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden;font-size:14px}',
  'th,td{padding:8px 10px;border-bottom:1px solid #eee;text-align:left;vertical-align:top}th{background:#2b2770;color:#fff;font-weight:700}',
  '.tw{overflow-x:auto}',
  '.pct{font-weight:900;padding:2px 8px;border-radius:10px}.g{background:#d9f7e3;color:#11692f}.y{background:#fff4cc;color:#7a5a00}.r{background:#ffe1e5;color:#9c1630}',
  '.q{background:#fff;border-radius:16px;padding:14px 16px;margin:10px 0;box-shadow:0 3px 0 #e2ddf7;border-left:6px solid #ccc}',
  '.q.ok1{border-left-color:#2fbf5b}.q.ok2{border-left-color:#ffb020}.q.bad{border-left-color:#e8394f}.q.wr{border-left-color:#3d7bff}',
  '.qh{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:6px}',
  '.tag{font-size:12px;font-weight:800;padding:3px 9px;border-radius:12px;background:#efeaff;color:#5b2fd6}',
  '.qt{font-size:17px;font-weight:700;line-height:1.45}',
  '.opts{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-top:10px}@media(max-width:600px){.opts{grid-template-columns:1fr}}',
  '.op{border:2px solid #e2ddf7;border-radius:10px;padding:7px 10px;font-weight:700;position:relative}',
  '.op.c{border-color:#2fbf5b;background:#e9fbef}.op.x{border-color:#e8394f;background:#fff0f2}',
  '.op small{float:right;font-weight:900;font-size:12px}.op.c small{color:#11692f}.op.x small{color:#9c1630}',
  '.ans{margin-top:8px;font-size:15px}.ans div{padding:6px 10px;border-radius:8px;margin:4px 0;background:#f6f5fb}',
  '.ans .ok{background:#e9fbef}.ans .no{background:#fff0f2}',
  '.essay{white-space:pre-wrap;background:#f6f8ff;border-radius:10px;padding:10px;font-size:15px;line-height:1.5}',
  'sup{font-size:.7em}o{text-decoration:overline}.sug{font-size:14px;color:#6b678f;margin-top:6px}',
  'input{font:inherit;font-size:18px;padding:10px 12px;border-radius:10px;border:2px solid #d9d3ff}button{font:inherit;font-weight:800;font-size:16px;padding:10px 16px;border-radius:10px;border:0;background:#ffcf33;cursor:pointer}'
].join('');

function pinPage_(base, tried) {
  return '<h1>Kết quả ôn tập của Quang Đạt</h1><p class="muted">Nhập mã xem kết quả.</p>' +
    (tried ? '<p style="color:#9c1630;font-weight:700">Sai mã, thử lại nhé.</p>' : '') +
    '<form method="get" action="' + esc_(base) + '"><input name="pin" type="password" placeholder="Mã" autofocus> <button>Xem</button></form>';
}

function readAll_() {
  const sh = ss_().getSheetByName(SH_JSON);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const out = [];
  v.forEach(function (r) {
    try { const d = JSON.parse(r.slice(1).join('')); out.push(d); } catch (e) { }
  });
  return out;
}
function readOne_(id) {
  const sh = ss_().getSheetByName(SH_JSON);
  if (!sh || sh.getLastRow() < 2) return null;
  const f = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  if (!f) return null;
  const r = sh.getRange(f.getRow(), 2, 1, sh.getLastColumn() - 1).getValues()[0];
  return JSON.parse(r.join(''));
}
function fmtTime_(iso) { try { return Utilities.formatDate(new Date(iso), TZ, 'dd/MM HH:mm'); } catch (e) { return ''; } }
function pc_(p) { return '<span class="pct ' + (p >= 90 ? 'g' : p >= 70 ? 'y' : 'r') + '">' + p + '%</span>'; }

function dashPage_(base, mon) {
  const all = readAll_().sort(function (a, b) { return String(b.finishedAt).localeCompare(String(a.finishedAt)); });
  const subj = {};
  const key = function (d) { return d.subjectKey + '~' + (d.student || ''); };
  all.forEach(function (d) { subj[key(d)] = d.subject.split('·')[0].trim() + ' – ' + (d.student || ''); });
  const keys = Object.keys(subj);
  if (!mon && keys.length) mon = keys[0];
  const list = all.filter(function (d) { return key(d) === mon; });
  const u = function (q) { return base + '?pin=' + encodeURIComponent(VIEW_PIN) + q; };
  const le = PropertiesService.getScriptProperties().getProperty('LAST_ERROR');
  let h = '<h1>Kết quả ôn tập</h1><p class="muted">Tổng cộng ' + all.length + ' lượt đã gửi. Bấm "Xem bài" để xem lại giao diện từng câu Đạt đã làm. <a href="' + esc_(ss_().getUrl()) + '">Mở Google Sheet</a></p>' +
    (le ? '<p style="color:#9c1630;font-size:13px">Lỗi gần nhất: ' + esc_(le) + '</p>' : '') + '<div class="tabs">';
  keys.forEach(function (k) { h += '<a class="' + (k === mon ? 'on' : '') + '" href="' + esc_(u('&mon=' + encodeURIComponent(k))) + '">' + esc_(subj[k]) + '</a>'; });
  h += '</div>';
  if (!list.length) return h + '<p>Chưa có lượt làm nào.</p>';
  // tổng hợp theo thử thách
  const lv = {};
  list.forEach(function (d) {
    const k = d.level;
    if (!lv[k]) lv[k] = { title: d.levelTitle, sub: d.levelSub, n: 0, best: 0, last: null, done: 0 };
    lv[k].n++; if (/^Hoàn thành/.test(d.status)) { lv[k].done++; lv[k].best = Math.max(lv[k].best, d.percent); }
    if (!lv[k].last) lv[k].last = d;
  });
  h += '<h2>Tiến độ từng thử thách</h2><div class="grid">';
  Object.keys(lv).sort(function (a, b) { return a - b; }).forEach(function (k) {
    const x = lv[k];
    h += '<div class="card"><span class="muted">Thử thách ' + k + '</span><br><b>' + esc_(x.title) + '</b><div class="muted">' + esc_(x.sub) + '</div>' +
      '<div class="big">' + x.best + '%</div><div class="muted">Tốt nhất · ' + x.n + ' lượt (' + x.done + ' hoàn thành) · lần cuối ' + x.last.percent + '%</div>' +
      '<div class="bar"><i style="width:' + x.best + '%"></i></div></div>';
  });
  h += '</div><h2>Tất cả các lượt</h2><div class="tw"><table><tr><th>Lúc</th><th>Thử thách</th><th>Trạng thái</th><th>Đúng ngay</th><th>%</th><th>Sao</th><th>Thời gian</th><th>Gợi ý</th><th></th></tr>';
  list.forEach(function (d) {
    h += '<tr><td>' + fmtTime_(d.finishedAt) + '</td><td>' + d.level + '. ' + esc_(d.levelTitle) + '</td><td>' + esc_(d.status) + '</td><td>' + d.correctFirst + '/' + d.total + '</td><td>' + pc_(d.percent) + '</td><td>' + d.stars +
      '</td><td>' + dur_(d.seconds) + '</td><td>' + d.hintsUsed + '</td><td><a href="' + esc_(u('&id=' + encodeURIComponent(d.id))) + '">Xem bài</a></td></tr>';
  });
  return h + '</table></div>';
}

function reviewPage_(base, id) {
  const d = readOne_(id);
  const back = '<a href="' + esc_(base + '?pin=' + encodeURIComponent(VIEW_PIN) + (d ? '&mon=' + encodeURIComponent(d.subjectKey + '~' + (d.student || '')) : '')) + '">← Về trang tổng hợp</a>';
  if (!d) return back + '<h1>Không tìm thấy lượt làm này</h1>';
  let h = back + '<h1>' + esc_(d.levelTitle) + '</h1><p class="muted">' + esc_(d.subject) + ' · Thử thách ' + d.level + ': ' + esc_(d.levelSub) + '</p>' +
    '<div class="grid"><div class="card"><span class="muted">Đúng ngay lần đầu</span><div class="big">' + d.correctFirst + '/' + d.total + '</div>' + pc_(d.percent) + '</div>' +
    '<div class="card"><span class="muted">Thời gian</span><div class="big" style="font-size:22px">' + dur_(d.seconds) + '</div><span class="muted">' + fmtTime_(d.startedAt) + ' → ' + fmtTime_(d.finishedAt) + '</span></div>' +
    '<div class="card"><span class="muted">Sao · vàng</span><div class="big">' + d.stars + ' sao</div><span class="muted">' + d.coins + ' vàng · hết tim ' + d.revives + ' lần</span></div>' +
    '<div class="card"><span class="muted">Trạng thái</span><div class="big" style="font-size:20px">' + esc_(d.status) + '</div><span class="muted">' + d.hintsUsed + ' câu nhờ chú Độ gợi ý</span></div></div>';
  h += '<h2>Bài làm chi tiết</h2>';
  d.details.forEach(function (x) {
    const cls = x.type === 'write' ? 'wr' : x.first ? 'ok1' : x.final ? 'ok2' : 'bad';
    const lab = x.type === 'write' ? 'Bài viết – thầy chấm' : x.first ? 'Đúng ngay lần đầu' : x.final ? 'Đúng ở lần 2' : 'Sai cả 2 lần';
    h += '<div class="q ' + cls + '"><div class="qh"><span class="tag">Câu ' + x.n + '</span><span class="tag">' + lab + '</span>' +
      (x.hints ? '<span class="tag" style="background:#fff4cc;color:#7a5a00">Mở gợi ý ' + x.hints + '/' + x.hintTotal + ' bước</span>' : '') + '</div>' +
      '<div class="qt">' + safe_(x.q) + '</div>';
    const ans = x.answers || [];
    if (x.type === 'mcq' && x.options) {
      const picked = ans.map(function (a) { return String(a).charAt(0); });
      const right = String(x.correct).charAt(0);
      h += '<div class="opts">' + x.options.map(function (o, i) {
        const L = 'ABCDEFGH'.charAt(i), k = picked.indexOf(L);
        const c = L === right ? 'c' : k >= 0 ? 'x' : '';
        const note = k >= 0 ? 'Đạt chọn lần ' + (k + 1) : (L === right ? 'Đáp án' : '');
        return '<div class="op ' + c + '">' + (note ? '<small>' + note + '</small>' : '') + L + '. ' + safe_(o) + '</div>';
      }).join('') + '</div>';
    } else if (x.type === 'write') {
      h += '<div class="essay">' + esc_(ans[ans.length - 1] || '') + '</div><div class="muted">' + wc_(ans[ans.length - 1]) + ' từ</div>';
    } else {
      h += '<div class="ans">' + ans.map(function (a, i) {
        const good = (i === ans.length - 1) && x.final;
        return '<div class="' + (good ? 'ok' : 'no') + '"><b>Lần ' + (i + 1) + ':</b> ' + safe_(a) + '</div>';
      }).join('') + '<div><b>Đáp án đúng:</b> ' + safe_(x.correct) + '</div></div>';
    }
    h += '</div>';
  });
  return h + '<p>' + back + '</p>';
}
function wc_(s) { return (String(s || '').trim().match(/[A-Za-zÀ-ỹ0-9']+/g) || []).length; }
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
/** chỉ giữ lại vài thẻ định dạng an toàn (b, u, i, sup, o, br, đoạn gợi ý) */
function safe_(h) {
  return esc_(h).replace(/&lt;(\/?)(b|u|i|sup|o|em)&gt;/g, '<$1$2>').replace(/&lt;br\s*\/?&gt;/g, '<br>')
    .replace(/&lt;div class=&quot;sug&quot;&gt;/g, '<div class="sug">').replace(/&lt;\/div&gt;/g, '</div>');
}
