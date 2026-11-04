/**
 * Nomzod Yo'li — Google Apps Script backend (1-bosqich: suhbat anketasi)
 *
 * Bu skript Google Sheets fayliga biriktiriladi (Kengaytmalar → Apps Script).
 * Mini App shu skriptga so'rov yuboradi, skript esa ruxsatni tekshirib,
 * ma'lumotni varaqlarga yozadi.
 *
 * Script Properties (Loyiha sozlamalari → Skript xususiyatlari):
 *   BOT_TOKEN  — BotFather bergan token (Telegram orqali kirishni tekshirish uchun)
 *   PIN        — brauzerdan kirish uchun PIN kod (kamida 6 raqam)
 */

var SHEETS = {
  NOMZODLAR: 'Nomzodlar',
  ANKETA: 'Anketa',
  ISH: 'Ish_tajribasi',
  GURUHLAR: 'Guruhlar',
  DAVOMAT: 'Davomat',
  TOPSHIRIQLAR: 'Topshiriqlar',
  IMTIHON: 'Imtihon',
  SOZLAMALAR: 'Sozlamalar'
};

var HEADERS = {
  Nomzodlar: ['ID', 'Suhbat sanasi', "F.I.Sh.", 'Telefon', "Tug'ilgan sana", 'Yosh', 'Vakansiya',
    "E'lon manbai", 'Holat', 'Guruh', 'Qoldirgan kunlar', 'Imtihon balli', 'HR', 'Yangilangan'],
  Anketa: ['ID', 'Suhbat sanasi', "F.I.Sh.", "Tug'ilgan sana", 'Yosh', 'Telefon', 'Millati', 'Oilaviy holati',
    'Viloyat', 'Shahar / tuman', 'Mahalla', "Ko'cha", 'Uy',
    "Ma'lumoti", "OTM nomi", 'Fakultet', 'Kurs',
    'Ish tajribasi (oy)', "Oldingi o'rtacha oylik (so'm)",
    'Kompyuter savodxonligi', 'Dasturlar', 'Rus tili', 'Ingliz tili', 'Boshqa til', 'Boshqa ko\'nikma',
    'Smena', 'Nima uchun ishga olishimiz kerak', "O'ziga ta'rif",
    'Vakansiya', "E'lon manbai", 'Muloqot (1-5)', 'Motivatsiya (1-5)', "Tashqi ko'rinish (1-5)",
    'HR izohi', 'Qaror', "Ma'lumot tasdiqlandi", 'HR'],
  Ish_tajribasi: ['Nomzod ID', "F.I.Sh.", 'Tashkilot', 'Lavozim', 'Boshlagan', 'Tugatgan', "Bo'shash sababi"],
  Guruhlar: ['Guruh', 'Boshlanish sanasi', 'Imtihon sanasi', 'Nomzodlar soni', 'Holat'],
  Davomat: ['Sana', 'Guruh', 'Kun', 'Nomzod ID', "F.I.Sh.", 'Holat', 'Baho (1-5)', 'Izoh'],
  Topshiriqlar: ['№', 'Topshiriq', 'Baholash mezoni', 'Maksimal ball', 'Faol'],
  Imtihon: ['Sana', 'Nomzod ID', "F.I.Sh.", 'Guruh', 'Ballar', 'Jami ball', 'Natija', 'HR izohi'],
  Sozlamalar: ['Kalit', 'Qiymat', 'Izoh']
};

var DEFAULT_SETTINGS = [
  ["O'tish bali", 80, "Imtihondan o'tish uchun kerakli minimal ball (100 dan)"],
  ['Chetlatish (qoldirilgan kun)', 2, 'Shuncha kun dars qoldirgan nomzod chetlatiladi. Kechikish hisoblanmaydi'],
  ['HR Telegram ID', '', "Ilovaga kira oladigan HR'ning Telegram ID raqami. Bir nechta bo'lsa vergul bilan"],
  ['Vakansiyalar', 'Operator, Sotuv menejeri, Kassir', 'Anketada tanlanadigan lavozimlar, vergul bilan'],
  ["E'lon manbalari", "Telegram guruh, Tanish orqali, Instagram, Boshqa", "Nomzod qayerdan kelgani, vergul bilan"],
  ['Smenalar', "Kunduzgi, Kechki, Tungi, Farqi yo'q", 'Anketada tanlanadigan smenalar, vergul bilan'],
  ['Millatlar', "O'zbek, Rus, Qozoq, Tojik, Qoraqalpoq, Boshqa", 'Anketada tanlanadigan millatlar, vergul bilan']
];

var STATUS = {
  NEW: "O'qishga taklif qilindi",
  RESERVE: 'Zaxirada',
  REJECTED: 'Rad etildi',
  STUDYING: "O'qishda",
  EXPELLED: 'Chetlatildi'
};
var ATT = { PRESENT: 'Keldi', LATE: 'Kechikdi', ABSENT: 'Kelmadi' };
var TRAINING_DAYS = 5;

// ---------------------------------------------------------------------------
// O'rnatish: bir marta ishga tushiriladi (Apps Script muharririda setup → Run)
// ---------------------------------------------------------------------------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var h = HEADERS[name];
    var existing = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (existing.join('') === '') {
      sh.getRange(1, 1, 1, h.length).setValues([h]);
    }
    sh.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#e2f4ea').setWrap(true);
    sh.setFrozenRows(1);
  });
  var set = ss.getSheetByName(SHEETS.SOZLAMALAR);
  if (set.getLastRow() < 2) {
    set.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    set.autoResizeColumns(1, 3);
  }
  // Telefon va ID ustunlari matn bo'lib qolsin (Sheets raqamga aylantirmasin)
  ss.getSheetByName(SHEETS.NOMZODLAR).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SHEETS.NOMZODLAR).getRange('D:D').setNumberFormat('@');
  ss.getSheetByName(SHEETS.ANKETA).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SHEETS.ANKETA).getRange('F:F').setNumberFormat('@');
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Лист1') || ss.getSheetByName('List1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  return 'Tayyor';
}

// ---------------------------------------------------------------------------
// HTTP kirish nuqtalari
// ---------------------------------------------------------------------------
function doGet() {
  return json_({ ok: true, app: "Nomzod Yo'li", version: 1 });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "So'rov noto'g'ri formatda" });
  }
  try {
    var user = authorize_(req.auth || {});
    var data = req.data || {};
    switch (req.action) {
      case 'config': return json_({ ok: true, user: user, config: getConfig_() });
      case 'checkPhone': return json_({ ok: true, matches: findByPhone_(data.phone) });
      case 'submitAnketa': return json_(submitAnketa_(data, user));
      case 'recent': return json_({ ok: true, items: recent_(20) });
      case 'groupCandidates': return json_({ ok: true, items: groupCandidates_() });
      case 'createGroup': return json_(createGroup_(data));
      case 'groups': return json_({ ok: true, items: listGroups_() });
      case 'getDay': return json_(getDay_(data));
      case 'saveDay': return json_(saveDay_(data));
      default: return json_({ ok: false, error: "Noma'lum amal" });
    }
  } catch (err) {
    var msg = err && err.userMessage ? err.userMessage : 'Serverda xatolik: ' + err;
    return json_({ ok: false, error: msg, code: err && err.code });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function userError_(message, code) {
  var e = new Error(message);
  e.userMessage = message;
  e.code = code || 'error';
  return e;
}

// ---------------------------------------------------------------------------
// Ruxsat: Telegram initData yoki PIN
// ---------------------------------------------------------------------------
function authorize_(auth) {
  var props = PropertiesService.getScriptProperties();
  if (auth.initData) {
    var token = props.getProperty('BOT_TOKEN');
    if (!token) throw userError_("Skriptda BOT_TOKEN sozlanmagan", 'config');
    var tg = verifyInitData_(auth.initData, token, 24 * 3600);
    if (!tg) throw userError_("Telegram ma'lumoti tasdiqlanmadi. Ilovani bot orqali qayta oching", 'auth');
    var allowed = String(getSetting_('HR Telegram ID') || '').split(/[,\s]+/).filter(String);
    if (allowed.indexOf(String(tg.id)) === -1) {
      throw userError_('Sizda ruxsat yo\'q. Telegram ID: ' + tg.id + '. Uni Sozlamalar varag\'iga qo\'shing', 'forbidden');
    }
    return { name: [tg.first_name, tg.last_name].filter(String).join(' '), id: tg.id, via: 'telegram' };
  }
  if (auth.pin) {
    var pin = props.getProperty('PIN');
    if (!pin) throw userError_('Skriptda PIN sozlanmagan', 'config');
    if (String(auth.pin) !== String(pin)) throw userError_("PIN noto'g'ri", 'auth');
    return { name: 'HR', id: null, via: 'pin' };
  }
  throw userError_('Kirish talab qilinadi', 'auth');
}

/**
 * Telegram Mini App initData imzosini tekshiradi.
 * https://core.telegram.org/bots/webapps#validating-data-received-via-the-mini-app
 * Qaytaradi: user obyekti yoki null.
 */
function verifyInitData_(initData, botToken, maxAgeSec) {
  var pairs = String(initData).split('&');
  var map = {};
  pairs.forEach(function (p) {
    var i = p.indexOf('=');
    if (i < 0) return;
    map[decodeURIComponent(p.slice(0, i))] = decodeURIComponent(p.slice(i + 1).replace(/\+/g, '%20'));
  });
  var hash = map.hash;
  if (!hash) return null;
  delete map.hash;
  var dcs = Object.keys(map).sort().map(function (k) { return k + '=' + map[k]; }).join('\n');
  var secret = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(botToken).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  var sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(dcs).getBytes(), secret);
  if (toHex_(sig) !== String(hash).toLowerCase()) return null;
  var authDate = Number(map.auth_date || 0);
  if (maxAgeSec && (Date.now() / 1000 - authDate) > maxAgeSec) return null;
  try { return JSON.parse(map.user); } catch (e) { return null; }
}

function toHex_(bytes) {
  return bytes.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');
}

// ---------------------------------------------------------------------------
// Sozlamalar
// ---------------------------------------------------------------------------
function settingsMap_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SOZLAMALAR);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] !== '') map[String(r[0]).trim()] = r[1];
  });
  return map;
}

function getSetting_(key) { return settingsMap_()[key]; }

function list_(v) {
  return String(v || '').split(',').map(function (s) { return s.trim(); }).filter(String);
}

function getConfig_() {
  var s = settingsMap_();
  return {
    vakansiyalar: list_(s['Vakansiyalar']),
    manbalar: list_(s["E'lon manbalari"]),
    smenalar: list_(s['Smenalar']),
    millatlar: list_(s['Millatlar'])
  };
}

// ---------------------------------------------------------------------------
// Anketa
// ---------------------------------------------------------------------------
function normPhone_(p) {
  var d = String(p || '').replace(/\D/g, '');
  if (d.length === 9) d = '998' + d;
  return d;
}

function fmtPhone_(p) {
  var d = normPhone_(p);
  if (d.length !== 12) return String(p || '');
  return '+' + d.slice(0, 3) + ' ' + d.slice(3, 5) + ' ' + d.slice(5, 8) + ' ' + d.slice(8, 10) + ' ' + d.slice(10, 12);
}

function findByPhone_(phone) {
  var target = normPhone_(phone);
  if (target.length < 12) return [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.NOMZODLAR);
  if (sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.Nomzodlar.length).getValues();
  return rows.filter(function (r) { return normPhone_(r[3]) === target; }).map(function (r) {
    return { id: r[0], date: fmtDate_(r[1]), name: r[2], vakansiya: r[6], holat: r[8] };
  });
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  return String(v || '');
}

function nextId_(now) {
  var props = PropertiesService.getScriptProperties();
  var prefix = 'N-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMM');
  var key = 'SEQ_' + prefix;
  var n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return prefix + '-' + ('000' + n).slice(-4);
}

function ageFrom_(isoDate, now) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate || ''));
  if (!m) return '';
  var y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
  var age = now.getFullYear() - y;
  if (now.getMonth() < mo || (now.getMonth() === mo && now.getDate() < d)) age--;
  return age;
}

function isoToDate_(iso) {
  var m = /^(\d{4})-(\d{2})(?:-(\d{2}))?$/.exec(String(iso || ''));
  if (!m) return '';
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3] || 1));
}

function monthsBetween_(from, to) {
  var a = /^(\d{4})-(\d{2})$/.exec(String(from || ''));
  if (!a) return 0;
  var b = /^(\d{4})-(\d{2})$/.exec(String(to || ''));
  var now = new Date();
  var ty = b ? Number(b[1]) : now.getFullYear();
  var tm = b ? Number(b[2]) : now.getMonth() + 1;
  return Math.max(0, (ty - Number(a[1])) * 12 + (tm - Number(a[2])) + 1);
}

function fmtMonth_(v) {
  var m = /^(\d{4})-(\d{2})$/.exec(String(v || ''));
  return m ? m[2] + '.' + m[1] : (v ? String(v) : 'Hozirgacha');
}

var REQUIRED = [
  ['fio', "F.I.Sh."], ['tugilgan', "Tug'ilgan sana"], ['telefon', 'Telefon'],
  ['viloyat', 'Viloyat'], ['malumot', "Ma'lumoti"], ['smena', 'Smena'], ['qaror', 'Qaror']
];

function submitAnketa_(d, user) {
  REQUIRED.forEach(function (f) {
    if (!d[f[0]] || String(d[f[0]]).trim() === '') throw userError_(f[1] + ' kiritilmagan', 'validation');
  });
  if (normPhone_(d.telefon).length !== 12) throw userError_("Telefon raqami to'liq emas", 'validation');
  if ([STATUS.NEW, STATUS.RESERVE, STATUS.REJECTED].indexOf(d.qaror) === -1) {
    throw userError_("Qaror noto'g'ri", 'validation');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Bir xil anketa ikki marta yuborilmasin (internet uzilib, qayta bosilganda)
    var props = PropertiesService.getScriptProperties();
    if (d.clientId) {
      var prev = props.getProperty('SUB_' + d.clientId);
      if (prev) return { ok: true, id: prev, duplicate: true };
    }

    var now = new Date();
    var id = nextId_(now);
    var hr = d.hrName || user.name || 'HR';
    var ish = Array.isArray(d.ish) ? d.ish : [];
    var months = ish.reduce(function (s, j) { return s + monthsBetween_(j.boshlagan, j.tugatgan); }, 0);
    var age = ageFrom_(d.tugilgan, now);
    var phone = fmtPhone_(d.telefon);
    var birth = isoToDate_(d.tugilgan);
    var dasturlar = (d.dasturlar || []).join(', ');

    var nRow = {};
    nRow['ID'] = id;
    nRow['Suhbat sanasi'] = now;
    nRow["F.I.Sh."] = d.fio;
    nRow['Telefon'] = phone;
    nRow["Tug'ilgan sana"] = birth;
    nRow['Yosh'] = age;
    nRow['Vakansiya'] = d.vakansiya || '';
    nRow["E'lon manbai"] = d.manba || '';
    nRow['Holat'] = d.qaror;
    nRow['Guruh'] = '';
    nRow['Qoldirgan kunlar'] = '';
    nRow['Imtihon balli'] = '';
    nRow['HR'] = hr;
    nRow['Yangilangan'] = now;
    appendObj_(ss.getSheetByName(SHEETS.NOMZODLAR), HEADERS.Nomzodlar, nRow);

    var a = {};
    a['ID'] = id; a['Suhbat sanasi'] = now; a["F.I.Sh."] = d.fio; a["Tug'ilgan sana"] = birth; a['Yosh'] = age;
    a['Telefon'] = phone; a['Millati'] = d.millati || ''; a['Oilaviy holati'] = d.oilaviy || '';
    a['Viloyat'] = d.viloyat || ''; a['Shahar / tuman'] = d.tuman || ''; a['Mahalla'] = d.mahalla || '';
    a["Ko'cha"] = d.kocha || ''; a['Uy'] = d.uy || '';
    a["Ma'lumoti"] = d.malumot || ''; a["OTM nomi"] = d.otm || ''; a['Fakultet'] = d.fakultet || ''; a['Kurs'] = d.kurs || '';
    a['Ish tajribasi (oy)'] = months; a["Oldingi o'rtacha oylik (so'm)"] = d.oylik ? Number(String(d.oylik).replace(/\D/g, '')) : '';
    a['Kompyuter savodxonligi'] = d.kompyuter || ''; a['Dasturlar'] = dasturlar;
    a['Rus tili'] = d.rus || ''; a['Ingliz tili'] = d.ingliz || ''; a['Boshqa til'] = d.boshqaTil || '';
    a["Boshqa ko'nikma"] = d.boshqa || ''; a['Smena'] = d.smena || '';
    a['Nima uchun ishga olishimiz kerak'] = d.nimaUchun || ''; a["O'ziga ta'rif"] = d.tarif || '';
    a['Vakansiya'] = d.vakansiya || ''; a["E'lon manbai"] = d.manba || '';
    a['Muloqot (1-5)'] = d.muloqot || ''; a['Motivatsiya (1-5)'] = d.motivatsiya || ''; a["Tashqi ko'rinish (1-5)"] = d.korinish || '';
    a['HR izohi'] = d.izoh || ''; a['Qaror'] = d.qaror; a["Ma'lumot tasdiqlandi"] = d.tasdiq ? 'Ha' : "Yo'q"; a['HR'] = hr;
    appendObj_(ss.getSheetByName(SHEETS.ANKETA), HEADERS.Anketa, a);

    var ishSh = ss.getSheetByName(SHEETS.ISH);
    ish.forEach(function (j) {
      if (!j.tashkilot) return;
      var r = {};
      r['Nomzod ID'] = id; r["F.I.Sh."] = d.fio; r['Tashkilot'] = j.tashkilot; r['Lavozim'] = j.lavozim || '';
      r['Boshlagan'] = fmtMonth_(j.boshlagan); r['Tugatgan'] = fmtMonth_(j.tugatgan); r["Bo'shash sababi"] = j.sabab || '';
      appendObj_(ishSh, HEADERS.Ish_tajribasi, r);
    });

    if (d.clientId) props.setProperty('SUB_' + d.clientId, id);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

/** "+998…", "-…", "=…" kabi matnni Sheets formula deb o'qimasligi uchun oldiga ' qo'yadi. */
function safe_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

/**
 * Bir marta ishga tushiring: telefon raqami #ERROR! bo'lib qolgan eski qatorlarni tiklaydi.
 */
function tuzatishTelefon() {
  var fixed = 0;
  [SHEETS.NOMZODLAR, SHEETS.ANKETA].forEach(function (name) {
    var sh = sheet_(name);
    if (sh.getLastRow() < 2) return;
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var c = headers.indexOf('Telefon');
    if (c < 0) return;
    var rng = sh.getRange(2, c + 1, sh.getLastRow() - 1, 1);
    var formulas = rng.getFormulas();
    formulas.forEach(function (f, i) {
      var text = String(f[0] || '');
      if (/^=\+?\d/.test(text)) {
        sh.getRange(i + 2, c + 1).setValue("'" + fmtPhone_(text));
        fixed++;
      }
    });
  });
  return fixed + ' ta raqam tuzatildi';
}

/** Sarlavhalar bo'yicha qator qo'shadi: ustunlar tartibi o'zgarsa ham to'g'ri joyga yozadi. */
function appendObj_(sh, defaultHeaders, obj) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.join('') === '') headers = defaultHeaders;
  var row = headers.map(function (h) { return obj.hasOwnProperty(h) ? safe_(obj[h]) : ''; });
  sh.appendRow(row);
}

function recent_(limit) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.NOMZODLAR);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - limit + 1);
  var rows = sh.getRange(start, 1, last - start + 1, HEADERS.Nomzodlar.length).getValues();
  return rows.reverse().map(function (r) {
    return { id: r[0], date: fmtDate_(r[1]), name: r[2], phone: r[3], vakansiya: r[6], holat: r[8] };
  });
}


// ---------------------------------------------------------------------------
// 2-bosqich: guruhlar va davomat
// ---------------------------------------------------------------------------
function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

/** Varaqni sarlavhalar bo'yicha o'qiydi: { headers, rows: [{_row, ...}] } */
function table_(name) {
  var sh = sheet_(name);
  var lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var rows = [];
  if (lastRow >= 2) {
    sh.getRange(2, 1, lastRow - 1, lastCol).getValues().forEach(function (r, i) {
      var o = { _row: i + 2 };
      headers.forEach(function (h, j) { if (h !== '') o[h] = r[j]; });
      rows.push(o);
    });
  }
  return { sh: sh, headers: headers, rows: rows };
}

function setCell_(t, rowNum, header, value) {
  var c = t.headers.indexOf(header);
  if (c > -1) t.sh.getRange(rowNum, c + 1).setValue(safe_(value));
}

function isoDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

function threshold_() { return Number(getSetting_('Chetlatish (qoldirilgan kun)')) || 2; }

function groupCandidates_() {
  return table_(SHEETS.NOMZODLAR).rows.filter(function (r) {
    return r['Holat'] === STATUS.NEW && !r['Guruh'];
  }).map(function (r) {
    return { id: r['ID'], name: r["F.I.Sh."], vakansiya: r['Vakansiya'], date: fmtDate_(r['Suhbat sanasi']) };
  });
}

function createGroup_(d) {
  var ids = Array.isArray(d.ids) ? d.ids.map(String) : [];
  if (!ids.length) throw userError_('Kamida bitta nomzodni tanlang', 'validation');
  var start = isoToDate_(d.start);
  if (!start) throw userError_("Boshlanish sanasini tanlang", 'validation');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var g = table_(SHEETS.GURUHLAR);
    var maxN = g.rows.reduce(function (m, r) { var n = Number(String(r['Guruh']).replace(/\D/g, '')); return n > m ? n : m; }, 0);
    var name = '#' + (maxN + 1);
    var exam = new Date(start.getTime()); exam.setDate(exam.getDate() + TRAINING_DAYS);
    var n = table_(SHEETS.NOMZODLAR);
    var picked = n.rows.filter(function (r) { return ids.indexOf(String(r['ID'])) > -1 && r['Holat'] === STATUS.NEW && !r['Guruh']; });
    if (!picked.length) throw userError_("Tanlangan nomzodlar allaqachon guruhga qo'shilgan", 'validation');
    var now = new Date();
    picked.forEach(function (r) {
      setCell_(n, r._row, 'Guruh', name);
      setCell_(n, r._row, 'Holat', STATUS.STUDYING);
      setCell_(n, r._row, 'Qoldirgan kunlar', 0);
      setCell_(n, r._row, 'Yangilangan', now);
    });
    var row = {};
    row['Guruh'] = name; row['Boshlanish sanasi'] = start; row['Imtihon sanasi'] = exam;
    row['Nomzodlar soni'] = picked.length; row['Holat'] = 'Faol';
    appendObj_(g.sh, HEADERS.Guruhlar, row);
    return { ok: true, group: name, count: picked.length, exam: isoDate_(exam) };
  } finally {
    lock.releaseLock();
  }
}

function listGroups_() {
  var members = table_(SHEETS.NOMZODLAR).rows;
  var att = table_(SHEETS.DAVOMAT).rows;
  return table_(SHEETS.GURUHLAR).rows.map(function (g) {
    var name = String(g['Guruh']);
    var mine = members.filter(function (m) { return String(m['Guruh']) === name; });
    var days = {};
    att.forEach(function (a) { if (String(a['Guruh']) === name) days[Number(a['Kun'])] = true; });
    return {
      group: name, start: isoDate_(g['Boshlanish sanasi']), exam: isoDate_(g['Imtihon sanasi']), holat: g['Holat'],
      total: mine.length,
      studying: mine.filter(function (m) { return m['Holat'] === STATUS.STUDYING; }).length,
      expelled: mine.filter(function (m) { return m['Holat'] === STATUS.EXPELLED; }).length,
      daysDone: Object.keys(days).map(Number).sort()
    };
  }).reverse();
}

function getDay_(d) {
  var group = String(d.group || ''), day = Number(d.day);
  if (!group) throw userError_('Guruh tanlanmagan', 'validation');
  if (!(day >= 1 && day <= TRAINING_DAYS)) throw userError_("Kun 1 dan " + TRAINING_DAYS + " gacha bo'lishi kerak", 'validation');
  var att = table_(SHEETS.DAVOMAT).rows.filter(function (a) { return String(a['Guruh']) === group; });
  var members = table_(SHEETS.NOMZODLAR).rows.filter(function (m) { return String(m['Guruh']) === group; });
  return {
    ok: true, group: group, day: day, threshold: threshold_(),
    items: members.map(function (m) {
      var id = String(m['ID']);
      var mineAll = att.filter(function (a) { return String(a['Nomzod ID']) === id; });
      var today = mineAll.filter(function (a) { return Number(a['Kun']) === day; })[0];
      return {
        id: id, name: m["F.I.Sh."], status: m['Holat'],
        absences: mineAll.filter(function (a) { return a['Holat'] === ATT.ABSENT && Number(a['Kun']) !== day; }).length,
        entry: today ? { holat: today['Holat'], baho: today['Baho (1-5)'], izoh: today['Izoh'] } : null
      };
    })
  };
}

function saveDay_(d) {
  var group = String(d.group || ''), day = Number(d.day);
  var entries = Array.isArray(d.entries) ? d.entries : [];
  if (!group || !(day >= 1 && day <= TRAINING_DAYS)) throw userError_("Guruh yoki kun noto'g'ri", 'validation');
  entries.forEach(function (e) {
    if ([ATT.PRESENT, ATT.LATE, ATT.ABSENT].indexOf(e.holat) === -1) throw userError_('Har bir nomzod uchun davomatni belgilang', 'validation');
  });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var warned = [], expelled = [], restored = [];
  try {
    var n = table_(SHEETS.NOMZODLAR);
    var members = {};
    n.rows.forEach(function (m) { if (String(m['Guruh']) === group) members[String(m['ID'])] = m; });
    var att = table_(SHEETS.DAVOMAT);
    var now = new Date();
    entries.forEach(function (e) {
      var m = members[String(e.id)];
      if (!m) return;
      var existing = att.rows.filter(function (a) { return String(a['Guruh']) === group && Number(a['Kun']) === day && String(a['Nomzod ID']) === String(e.id); })[0];
      var baho = e.holat === ATT.ABSENT ? '' : (e.baho ? Number(e.baho) : '');
      if (existing) {
        setCell_(att, existing._row, 'Holat', e.holat);
        setCell_(att, existing._row, 'Baho (1-5)', baho);
        setCell_(att, existing._row, 'Izoh', e.izoh || '');
        existing['Holat'] = e.holat;
      } else {
        var row = {};
        row['Sana'] = now; row['Guruh'] = group; row['Kun'] = day; row['Nomzod ID'] = String(e.id);
        row["F.I.Sh."] = m["F.I.Sh."]; row['Holat'] = e.holat; row['Baho (1-5)'] = baho; row['Izoh'] = e.izoh || '';
        appendObj_(att.sh, HEADERS.Davomat, row);
        att.rows.push({ 'Guruh': group, 'Kun': day, 'Nomzod ID': String(e.id), 'Holat': e.holat });
      }
    });
    var limit = threshold_();
    Object.keys(members).forEach(function (id) {
      var m = members[id];
      var before = Number(m['Qoldirgan kunlar']) || 0;
      var count = att.rows.filter(function (a) { return String(a['Guruh']) === group && String(a['Nomzod ID']) === id && a['Holat'] === ATT.ABSENT; }).length;
      if (count !== before) setCell_(n, m._row, 'Qoldirgan kunlar', count);
      if (count >= limit && m['Holat'] === STATUS.STUDYING) {
        setCell_(n, m._row, 'Holat', STATUS.EXPELLED); setCell_(n, m._row, 'Yangilangan', now);
        expelled.push(m["F.I.Sh."]);
      } else if (count < limit && m['Holat'] === STATUS.EXPELLED) {
        setCell_(n, m._row, 'Holat', STATUS.STUDYING); setCell_(n, m._row, 'Yangilangan', now);
        restored.push(m["F.I.Sh."]);
      } else if (count > before && count < limit) {
        warned.push(m["F.I.Sh."]);
      }
    });
  } finally {
    lock.releaseLock();
  }
  var lines = [];
  if (warned.length) lines.push('⚠️ Guruh ' + group + ', ' + day + '-kun: dars qoldirdi (yana bir marta qoldirsa chetlatiladi):\n• ' + warned.join('\n• '));
  if (expelled.length) lines.push('⛔ Guruh ' + group + ', ' + day + '-kun: chetlatildi (' + limit + ' kun qoldirdi):\n• ' + expelled.join('\n• '));
  if (lines.length) notify_(lines.join('\n\n'));
  return { ok: true, warned: warned, expelled: expelled, restored: restored };
}

/** HR'ga bot orqali xabar yuboradi. HR botga kamida bir marta /start yozgan bo'lishi kerak. */
function notify_(text) {
  var token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) return;
  var ids = String(getSetting_('HR Telegram ID') || '').split(/[,\s]+/).filter(String);
  ids.forEach(function (id) {
    try {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ chat_id: id, text: text })
      });
    } catch (e) { /* xabar ketmasa ham davomat saqlanadi */ }
  });
}

/** Bir marta qo'lda ishga tushiring: Telegram'ga xabar yuborish ruxsatini beradi va sinab ko'radi. */
function testXabar() {
  notify_("✅ Nomzod Yo'li: bot xabarlari ishlayapti.");
  return 'Yuborildi';
}
