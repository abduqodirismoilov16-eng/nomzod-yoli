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
 *   OQUV_BOT_TOKEN — o'quv botining tokeni (3-bosqich: qo'llanmalar)
 */

var SHEETS = {
  NOMZODLAR: 'Nomzodlar',
  ANKETA: 'Anketa',
  ISH: 'Ish_tajribasi',
  GURUHLAR: 'Guruhlar',
  DAVOMAT: 'Davomat',
  TOPSHIRIQLAR: 'Topshiriqlar',
  IMTIHON: 'Imtihon',
  SOZLAMALAR: 'Sozlamalar',
  QOLLANMALAR: "Qo'llanmalar",
  YUBORISHLAR: 'Yuborishlar',
  SAVOLLAR: 'Savollar',
  ALOQA: 'Aloqa'
};

var HEADERS = {
  Nomzodlar: ['ID', 'Suhbat sanasi', "F.I.Sh.", 'Telefon', "Tug'ilgan sana", 'Yosh', 'Vakansiya',
    "E'lon manbai", 'Holat', 'Guruh', 'Qoldirgan kunlar', 'Imtihon balli', 'HR', 'Yangilangan', 'Telegram ID',
    'Oxirgi aloqa', 'Aloqa usuli', 'Aloqa natijasi', "O'qish kuni", 'Boshlagan sana', 'Kelmaslik sababi'],
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
  Sozlamalar: ['Kalit', 'Qiymat', 'Izoh'],
  "Qo'llanmalar": ['Kun', 'Tartib', 'Nomi', 'Turi', 'Fayl ID', 'Faol', "Qo'shilgan"],
  Yuborishlar: ['Sana', 'Guruh', 'Kun', 'Nomzod ID', "F.I.Sh.", 'Holat', "O'qidi", 'Ochilgan'],
  Savollar: ['Sana', 'Nomzod ID', "F.I.Sh.", 'Guruh', 'Savol', 'Javob', 'Javob sanasi'],
  Aloqa: ['Sana', 'Nomzod ID', "F.I.Sh.", 'Usul', 'Natija', "O'qish kuni", 'Izoh', 'HR']
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
  EXPELLED: 'Chetlatildi',
  NOSHOW: "O'qishga kelmadi"
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
    } else {
      // Keyingi bosqichlarda qo'shilgan ustunlarni oxiriga qo'shadi, mavjudlariga tegmaydi
      var missing = h.filter(function (x) { return existing.indexOf(x) === -1; });
      if (missing.length) sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing])
        .setFontWeight('bold').setBackground('#e2f4ea').setWrap(true);
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
  // O'quv botidan kelgan Telegram xabarlari (webhook). Hech narsa qaytarmaymiz: Telegram 200 javobini oladi.
  if (e && e.parameter && e.parameter.bot) { oquvWebhook_(e); return; }
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "So'rov noto'g'ri formatda" });
  }
  try {
    // Nomzodlar ilovasi («Darslarim»): o'quv boti imzosi bilan kiradi
    if (String(req.action || '').indexOf('oq.') === 0) return json_(oquvApi_(req));
    var user = authorize_(req.auth || {});
    var data = req.data || {};
    switch (req.action) {
      case 'config': {
        // Ilova ochilganda bitta so'rov: sozlamalar + ro'yxat birga
        var cfgOut = { ok: true, user: user, config: getConfig_() };
        if (data.recent) { ensureSchema_(); cfgOut.recent = recent_(data.recent); }
        return json_(cfgOut);
      }
      case 'checkPhone': return json_({ ok: true, matches: findByPhone_(data.phone) });
      case 'submitAnketa': return json_(submitAnketa_(data, user));
      case 'recent': ensureSchema_(); return json_(recent_(data.filter));
      case 'addContact': return json_(addContact_(data, user));
      case 'setNoShow': return json_(setNoShow_(data, user));
      case 'groupCandidates': return json_({ ok: true, items: groupCandidates_() });
      case 'createGroup': return json_(createGroup_(data));
      case 'addToGroup': return json_(addToGroup_(data));
      case 'groups': ensureSchema_(); return json_({ ok: true, items: listGroups_() });
      case 'getDay': return json_(getDay_(data));
      case 'saveDay': ensureSchema_(); return json_(saveDay_(data));
      case 'getCandidate': ensureSchema_(); return json_(getCandidate_(data));
      case 'updateAnketa': return json_(updateAnketa_(data, user));
      case 'search': return json_({ ok: true, items: search_(data.q) });
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
    if (allowed.indexOf(String(tg.id)) === -1) {   // yangi qo'shilgan bo'lishi mumkin: keshsiz qayta tekshiramiz
      keshniTozalash();
      allowed = String(getSetting_('HR Telegram ID') || '').split(/[,\s]+/).filter(String);
    }
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
/**
 * Sozlamalar 5 daqiqaga keshda saqlanadi (har so'rovda varaqni o'qimaslik uchun).
 * Sozlamalar varag'ini qo'lda o'zgartirsangiz, kuchga kirishi 5 daqiqagacha cho'zilishi mumkin
 * yoki keshniTozalash() ni ishga tushiring.
 */
var SETTINGS_CACHE_KEY = 'sozlamalar_v1', settingsMemo_ = null;
function settingsMap_() {
  if (settingsMemo_) return settingsMemo_;
  var cache = CacheService.getScriptCache();
  var hit = cache.get(SETTINGS_CACHE_KEY);
  if (hit) { try { settingsMemo_ = JSON.parse(hit); return settingsMemo_; } catch (e) {} }
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SOZLAMALAR);
  var map = {};
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0] !== '') map[String(r[0]).trim()] = r[1];
    });
  }
  try { cache.put(SETTINGS_CACHE_KEY, JSON.stringify(map), 300); } catch (e) {}
  settingsMemo_ = map;
  return map;
}

/** Sozlamalar varag'i qo'lda o'zgartirilsa kesh avtomatik tozalanadi */
function onEdit(e) {
  try { if (e && e.range && e.range.getSheet().getName() === SHEETS.SOZLAMALAR) keshniTozalash(); } catch (err) {}
}

/** Sozlamalar varag'ini o'zgartirgandan keyin darhol kuchga kiritish uchun */
function keshniTozalash() {
  CacheService.getScriptCache().remove(SETTINGS_CACHE_KEY);
  settingsMemo_ = null;
  return 'Kesh tozalandi';
}

/**
 * Bir marta ishga tushiring: skript har 5 daqiqada "uyg'onib" turadi,
 * shunda ilova birinchi ochilganda uzoq kutmaydi.
 */
function tezlikniYoqish() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'uygoqSaqlash') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('uygoqSaqlash').timeBased().everyMinutes(5).create();
  uygoqSaqlash();
  return 'Tayyor: har 5 daqiqada uyg\'otib turiladi';
}

function uygoqSaqlash() {
  settingsMemo_ = null;
  CacheService.getScriptCache().remove(SETTINGS_CACHE_KEY);
  settingsMap_();   // keshni yangilaydi va Sheets ulanishini "isitadi"
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
    var values = rng.getDisplayValues();
    formulas.forEach(function (f, i) {
      var text = String(f[0] || values[i][0] || '');
      if (/^=?\+?[\d ]{9,}$/.test(text.trim()) && normPhone_(text).length === 12 && String(values[i][0]).indexOf('#') === 0) {
        sh.getRange(i + 2, c + 1).setValue("'" + fmtPhone_(text));
        fixed++;
      }
    });
  });
  Logger.log(fixed + ' ta raqam tuzatildi');
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

var FILTERS = ['all', 'noContact', 'waiting', 'studying', 'noshow'];

/** Ro'yxat qatori: holat + aloqa/boshlash belgilari uchun qisqa ma'lumot */
function listItem_(r) {
  return {
    id: r['ID'], date: fmtDate_(r['Suhbat sanasi']), name: r["F.I.Sh."], phone: r['Telefon'], vakansiya: r['Vakansiya'], holat: r['Holat'],
    guruh: r['Guruh'] || '', aloqa: r['Oxirgi aloqa'] ? shortDate_(r['Oxirgi aloqa']) : '', usul: r['Aloqa usuli'] || '',
    natija: r['Aloqa natijasi'] || '', boshlagan: r['Boshlagan sana'] ? shortDate_(r['Boshlagan sana']) : ''
  };
}

/** Qaysi filtrga tushadi */
function filterOf_(r) {
  var h = r['Holat'];
  if (h === STATUS.NOSHOW) return 'noshow';
  if (h === STATUS.NEW) return r['Aloqa natijasi'] === CONTACT_RES.YES || r['Aloqa natijasi'] === CONTACT_RES.WAIT ? 'waiting' : 'noContact';
  if (h === STATUS.STUDYING) return r['Boshlagan sana'] ? 'studying' : 'waiting';
  return '';
}

function recent_(filter) {
  filter = FILTERS.indexOf(filter) > -1 ? filter : 'all';
  var rows = table_(SHEETS.NOMZODLAR).rows.filter(function (r) { return r['ID'] !== ''; });
  var counts = { noContact: 0, waiting: 0, studying: 0, noshow: 0 };
  rows.forEach(function (r) { var f = filterOf_(r); if (f) counts[f]++; });
  var picked = filter === 'all' ? rows.slice(-20) : rows.filter(function (r) { return filterOf_(r) === filter; }).slice(-100);
  return { ok: true, filter: filter, counts: counts, items: picked.reverse().map(listItem_) };
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

/** Mavjud guruhga yangi nomzodlarni qo'shadi (o'qish boshlangandan keyin kelganlar uchun). */
function addToGroup_(d) {
  var group = String(d.group || '');
  var ids = Array.isArray(d.ids) ? d.ids.map(String) : [];
  if (!ids.length) throw userError_('Kamida bitta nomzodni tanlang', 'validation');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var g = table_(SHEETS.GURUHLAR);
    var gr = g.rows.filter(function (r) { return String(r['Guruh']) === group; })[0];
    if (!gr) throw userError_('Guruh topilmadi', 'validation');
    var n = table_(SHEETS.NOMZODLAR);
    var picked = n.rows.filter(function (r) { return ids.indexOf(String(r['ID'])) > -1 && r['Holat'] === STATUS.NEW && !r['Guruh']; });
    if (!picked.length) throw userError_("Tanlangan nomzodlar allaqachon guruhga qo'shilgan", 'validation');
    var now = new Date();
    picked.forEach(function (r) {
      setCell_(n, r._row, 'Guruh', group);
      setCell_(n, r._row, 'Holat', STATUS.STUDYING);
      setCell_(n, r._row, 'Qoldirgan kunlar', 0);
      setCell_(n, r._row, 'Yangilangan', now);
    });
    var total = n.rows.filter(function (r) { return String(r['Guruh']) === group; }).length + picked.length;
    setCell_(g, gr._row, 'Nomzodlar soni', total);
    return { ok: true, group: group, count: picked.length };
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
      started: mine.filter(function (m) { return m['Holat'] === STATUS.STUDYING && m['Boshlagan sana']; }).length,
      notStarted: mine.filter(function (m) { return m['Holat'] === STATUS.STUDYING && !m['Boshlagan sana']; }).length,
      noshow: mine.filter(function (m) { return m['Holat'] === STATUS.NOSHOW; }).length,
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
        id: id, name: m["F.I.Sh."], status: m['Holat'], started: !!m['Boshlagan sana'],
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
  var warned = [], expelled = [], restored = [], noshow = [];
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
    var gr = groupRow_(group);
    var TRACKED = [STATUS.STUDYING, STATUS.EXPELLED, STATUS.NOSHOW];
    Object.keys(members).forEach(function (id) {
      var m = members[id];
      var before = Number(m['Qoldirgan kunlar']) || 0;
      var count = att.rows.filter(function (a) { return String(a['Guruh']) === group && String(a['Nomzod ID']) === id && a['Holat'] === ATT.ABSENT; }).length;
      if (count !== before) setCell_(n, m._row, 'Qoldirgan kunlar', count);
      // Boshlagan sana: birinchi "Keldi"/"Kechikdi" kuni
      var first = firstPresent_(att.rows, group, id);
      var startDate = first ? dayDate_(gr && gr['Boshlanish sanasi'], first) : '';
      var oldStart = m['Boshlagan sana'] instanceof Date ? m['Boshlagan sana'].getTime() : String(m['Boshlagan sana'] || '');
      if ((startDate instanceof Date ? startDate.getTime() : '') !== oldStart) setCell_(n, m._row, 'Boshlagan sana', startDate);
      var cur = m['Holat'];
      if (TRACKED.indexOf(cur) === -1) return;
      var target = cur;
      if (first) target = count >= limit ? STATUS.EXPELLED : STATUS.STUDYING;   // o'qishni boshlagan
      else if (count >= limit) target = STATUS.NOSHOW;                            // bir kun ham kelmagan
      else if (cur === STATUS.EXPELLED) target = STATUS.STUDYING;
      if (target !== cur) {
        setCell_(n, m._row, 'Holat', target); setCell_(n, m._row, 'Yangilangan', now);
        if (target === STATUS.NOSHOW) setCell_(n, m._row, 'Kelmaslik sababi', count + ' kun darsga kelmadi');
        else if (cur === STATUS.NOSHOW) setCell_(n, m._row, 'Kelmaslik sababi', '');
        if (target === STATUS.EXPELLED) expelled.push(m["F.I.Sh."]);
        else if (target === STATUS.NOSHOW) noshow.push(m["F.I.Sh."]);
        else restored.push(m["F.I.Sh."]);
      } else if (target === STATUS.STUDYING && count > before && count < limit) {
        warned.push(m["F.I.Sh."]);
      }
    });
  } finally {
    lock.releaseLock();
  }
  var lines = [];
  if (warned.length) lines.push('⚠️ Guruh ' + group + ', ' + day + '-kun: dars qoldirdi (yana bir marta qoldirsa chetlatiladi):\n• ' + warned.join('\n• '));
  if (expelled.length) lines.push('⛔ Guruh ' + group + ', ' + day + '-kun: chetlatildi (' + limit + ' kun qoldirdi):\n• ' + expelled.join('\n• '));
  if (noshow.length) lines.push('🚫 Guruh ' + group + ', ' + day + '-kun: o\'qishga kelmadi (bir kun ham kelmadi):\n• ' + noshow.join('\n• '));
  if (lines.length) notify_(lines.join('\n\n'));
  return { ok: true, warned: warned, expelled: expelled, restored: restored, noshow: noshow };
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


// ---------------------------------------------------------------------------
// Nomzod kartasi: ko'rish, tahrirlash, qidirish
// ---------------------------------------------------------------------------
// Ilova maydoni ↔ Anketa ustuni (oddiy matn maydonlari)
var ANKETA_FIELDS = [
  ['fio', "F.I.Sh."], ['millati', 'Millati'], ['oilaviy', 'Oilaviy holati'],
  ['viloyat', 'Viloyat'], ['tuman', 'Shahar / tuman'], ['mahalla', 'Mahalla'], ['kocha', "Ko'cha"], ['uy', 'Uy'],
  ['malumot', "Ma'lumoti"], ['otm', 'OTM nomi'], ['fakultet', 'Fakultet'], ['kurs', 'Kurs'],
  ['kompyuter', 'Kompyuter savodxonligi'], ['rus', 'Rus tili'], ['ingliz', 'Ingliz tili'],
  ['boshqaTil', 'Boshqa til'], ['boshqa', "Boshqa ko'nikma"], ['smena', 'Smena'],
  ['nimaUchun', 'Nima uchun ishga olishimiz kerak'], ['tarif', "O'ziga ta'rif"],
  ['vakansiya', 'Vakansiya'], ['manba', "E'lon manbai"],
  ['muloqot', 'Muloqot (1-5)'], ['motivatsiya', 'Motivatsiya (1-5)'], ['korinish', "Tashqi ko'rinish (1-5)"],
  ['izoh', 'HR izohi']
];
var PRE_GROUP = [STATUS.NEW, STATUS.RESERVE, STATUS.REJECTED];

function monthToIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  var m = /^(\d{2})\.(\d{4})$/.exec(String(v || '').trim());
  return m ? m[2] + '-' + m[1] : '';
}

function findRow_(t, id) {
  return t.rows.filter(function (r) { return String(r['ID'] || r['Nomzod ID']) === String(id); })[0];
}

function getCandidate_(d) {
  var id = String(d.id || '');
  var a = findRow_(table_(SHEETS.ANKETA), id);
  var n = findRow_(table_(SHEETS.NOMZODLAR), id);
  if (!a || !n) throw userError_('Nomzod topilmadi: ' + id, 'notfound');
  var out = {};
  ANKETA_FIELDS.forEach(function (f) { var v = a[f[1]]; out[f[0]] = v === undefined || v === null ? '' : String(v); });
  out.tugilgan = isoDate_(a["Tug'ilgan sana"]);
  out.telefon = fmtPhone_(a['Telefon'] || n['Telefon']);
  out.oylik = a["Oldingi o'rtacha oylik (so'm)"] ? String(a["Oldingi o'rtacha oylik (so'm)"]) : '';
  out.dasturlar = String(a['Dasturlar'] || '').split(',').map(function (x) { return x.trim(); }).filter(String);
  out.qaror = String(a['Qaror'] || '');
  out.tasdiq = a["Ma'lumot tasdiqlandi"] === 'Ha';
  out.hrName = String(a['HR'] || '');
  out.ish = table_(SHEETS.ISH).rows.filter(function (r) { return String(r['Nomzod ID']) === id; }).map(function (r) {
    return { tashkilot: String(r['Tashkilot'] || ''), lavozim: String(r['Lavozim'] || ''), boshlagan: monthToIso_(r['Boshlagan']),
      tugatgan: monthToIso_(r['Tugatgan']), sabab: String(r["Bo'shash sababi"] || '') };
  });
  return {
    ok: true, id: id, data: out,
    meta: { holat: n['Holat'], guruh: n['Guruh'] || '', qoldirgan: n['Qoldirgan kunlar'], imtihon: n['Imtihon balli'],
      suhbat: fmtDate_(n['Suhbat sanasi']), yangilangan: fmtDate_(n['Yangilangan']), hr: n['HR'],
      canChangeDecision: PRE_GROUP.indexOf(n['Holat']) > -1,
      aloqa: contactsOf_(id), oqishKuni: n["O'qish kuni"] ? fmtDate_(n["O'qish kuni"]) : '',
      boshlagan: n['Boshlagan sana'] ? fmtDate_(n['Boshlagan sana']) : '', sabab: n['Kelmaslik sababi'] || '',
      natija: n['Aloqa natijasi'] || '' }
  };
}

function updateAnketa_(d, user) {
  var id = String(d.id || '');
  var data = d.data || {};
  REQUIRED.forEach(function (f) {
    if (f[0] === 'qaror') return;
    if (!data[f[0]] || String(data[f[0]]).trim() === '') throw userError_(f[1] + ' kiritilmagan', 'validation');
  });
  if (normPhone_(data.telefon).length !== 12) throw userError_("Telefon raqami to'liq emas", 'validation');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var at = table_(SHEETS.ANKETA), nt = table_(SHEETS.NOMZODLAR);
    var a = findRow_(at, id), n = findRow_(nt, id);
    if (!a || !n) throw userError_('Nomzod topilmadi: ' + id, 'notfound');
    var now = new Date();
    var ish = Array.isArray(data.ish) ? data.ish.filter(function (j) { return j && j.tashkilot; }) : [];
    var months = ish.reduce(function (s, j) { return s + monthsBetween_(j.boshlagan, j.tugatgan); }, 0);
    var birth = isoToDate_(data.tugilgan);
    var phone = fmtPhone_(data.telefon);
    var age = ageFrom_(data.tugilgan, now);
    var holat = n['Holat'];
    var canChange = PRE_GROUP.indexOf(holat) > -1;
    var qaror = canChange && PRE_GROUP.indexOf(data.qaror) > -1 ? data.qaror : a['Qaror'];

    var aVals = {};
    ANKETA_FIELDS.forEach(function (f) { aVals[f[1]] = data[f[0]] == null ? '' : String(data[f[0]]); });
    aVals["Tug'ilgan sana"] = birth; aVals['Yosh'] = age; aVals['Telefon'] = phone;
    aVals['Ish tajribasi (oy)'] = months;
    aVals["Oldingi o'rtacha oylik (so'm)"] = data.oylik ? Number(String(data.oylik).replace(/\D/g, '')) : '';
    aVals['Dasturlar'] = (data.dasturlar || []).join(', ');
    aVals['Qaror'] = qaror;
    rowUpdate_(at, a, aVals);

    var nVals = {};
    nVals["F.I.Sh."] = data.fio; nVals['Telefon'] = phone; nVals["Tug'ilgan sana"] = birth; nVals['Yosh'] = age;
    nVals['Vakansiya'] = data.vakansiya || ''; nVals["E'lon manbai"] = data.manba || ''; nVals['Yangilangan'] = now;
    if (canChange) nVals['Holat'] = qaror;
    rowUpdate_(nt, n, nVals);

    // Ish tajribasi: eski qatorlarni o'chirib, yangisini yozamiz
    var it = table_(SHEETS.ISH);
    it.rows.filter(function (r) { return String(r['Nomzod ID']) === id; })
      .map(function (r) { return r._row; }).sort(function (x, y) { return y - x; })
      .forEach(function (rowNum) { it.sh.deleteRow(rowNum); });
    ish.forEach(function (j) {
      var r = {};
      r['Nomzod ID'] = id; r["F.I.Sh."] = data.fio; r['Tashkilot'] = j.tashkilot; r['Lavozim'] = j.lavozim || '';
      r['Boshlagan'] = fmtMonth_(j.boshlagan); r['Tugatgan'] = fmtMonth_(j.tugatgan); r["Bo'shash sababi"] = j.sabab || '';
      appendObj_(it.sh, HEADERS.Ish_tajribasi, r);
    });
    // Davomat varag'idagi ism ham yangilansin
    if (String(data.fio) !== String(n["F.I.Sh."])) {
      var dv = table_(SHEETS.DAVOMAT);
      dv.rows.forEach(function (r) { if (String(r['Nomzod ID']) === id) setCell_(dv, r._row, "F.I.Sh.", data.fio); });
    }
    return { ok: true, id: id, holat: canChange ? qaror : holat, decisionLocked: !canChange && data.qaror && data.qaror !== a['Qaror'] };
  } finally {
    lock.releaseLock();
  }
}

/** Faqat o'zgargan kataklarni yozadi */
function rowUpdate_(t, row, vals) {
  Object.keys(vals).forEach(function (h) {
    var oldV = row[h], newV = vals[h];
    var same = (oldV instanceof Date && newV instanceof Date) ? oldV.getTime() === newV.getTime() : String(oldV) === String(newV);
    if (!same) setCell_(t, row._row, h, newV);
  });
}

function search_(q) {
  q = String(q || '').trim().toLowerCase();
  if (q.length < 2) return [];
  var digits = q.replace(/\D/g, '');
  return table_(SHEETS.NOMZODLAR).rows.filter(function (r) {
    var name = String(r["F.I.Sh."] || '').toLowerCase();
    var id = String(r['ID'] || '').toLowerCase();
    var ph = String(r['Telefon'] || '').replace(/\D/g, '');
    return name.indexOf(q) > -1 || id.indexOf(q) > -1 || (digits.length >= 3 && ph.indexOf(digits) > -1);
  }).reverse().slice(0, 30).map(listItem_);
}


// ---------------------------------------------------------------------------
// Nomzodga xabar berish (qo'ng'iroq/SMS) va o'qishni boshlagan/kelmaganini belgilash
// ---------------------------------------------------------------------------
var CONTACT_WAYS = ["Qo'ng'iroq", 'SMS', 'Telegram'];
var CONTACT_RES = { YES: 'Keladi', WAIT: 'Kutilmoqda', NOANSWER: 'Javob bermadi', OFF: "Raqam o'chiq", NO: 'Rad etdi' };
var SCHEMA_VERSION = '2';

/**
 * Yangi ustun va "Aloqa" varag'ini avtomatik qo'shadi (setup'ni qayta ishga tushirish shart emas).
 * Birinchi marta ishlaganda mavjud guruhlar uchun "Boshlagan sana" ni davomatdan to'ldiradi va
 * bir kun ham kelmay chetlatilganlarni "O'qishga kelmadi" holatiga o'tkazadi.
 */
function ensureSchema_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA_VERSION') === SCHEMA_VERSION) return;
  withLock_(function () {
    if (props.getProperty('SCHEMA_VERSION') === SCHEMA_VERSION) return;
    ensureSheet_(SHEETS.ALOQA);
    ['Telegram ID', 'Oxirgi aloqa', 'Aloqa usuli', 'Aloqa natijasi', "O'qish kuni", 'Boshlagan sana', 'Kelmaslik sababi']
      .forEach(function (h) { ensureColumn_(SHEETS.NOMZODLAR, h); });
    var n = table_(SHEETS.NOMZODLAR), att = table_(SHEETS.DAVOMAT).rows;
    var starts = {};
    table_(SHEETS.GURUHLAR).rows.forEach(function (g) { starts[String(g['Guruh'])] = g['Boshlanish sanasi']; });
    n.rows.forEach(function (m) {
      if (!m['Guruh']) return;
      var first = firstPresent_(att, m['Guruh'], m['ID']);
      if (first && !m['Boshlagan sana']) setCell_(n, m._row, 'Boshlagan sana', dayDate_(starts[String(m['Guruh'])], first));
      if (!first && m['Holat'] === STATUS.EXPELLED) setCell_(n, m._row, 'Holat', STATUS.NOSHOW);
    });
    props.setProperty('SCHEMA_VERSION', SCHEMA_VERSION);
  });
}

/** Nomzod guruhda birinchi kelgan (Keldi/Kechikdi) kun raqami, kelmagan bo'lsa 0 */
function firstPresent_(attRows, group, id) {
  var first = 0;
  attRows.forEach(function (a) {
    if (String(a['Guruh']) !== String(group) || String(a['Nomzod ID']) !== String(id)) return;
    if (a['Holat'] !== ATT.PRESENT && a['Holat'] !== ATT.LATE) return;
    var d = Number(a['Kun']); if (!first || d < first) first = d;
  });
  return first;
}

/** Guruh boshlangan sanadan N-kun sanasi */
function dayDate_(start, day) {
  var s = isoToDate_(isoDate_(start));
  if (!s) return new Date();
  s.setDate(s.getDate() + day - 1);
  return s;
}

function contactsOf_(id) {
  var sh = sheet_(SHEETS.ALOQA);
  if (!sh) return [];
  return table_(SHEETS.ALOQA).rows.filter(function (r) { return String(r['Nomzod ID']) === String(id); }).map(function (r) {
    var t = r['Sana'];
    return { sana: t instanceof Date ? Utilities.formatDate(t, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : String(t || ''),
      usul: r['Usul'], natija: r['Natija'], oqish: r["O'qish kuni"] ? fmtDate_(r["O'qish kuni"]) : '', izoh: r['Izoh'] || '', hr: r['HR'] || '' };
  });
}

var CONTACT_ALLOWED = [STATUS.NEW, STATUS.NOSHOW];

function addContact_(d, user) {
  var id = String(d.id || '');
  if (CONTACT_WAYS.indexOf(d.usul) === -1) throw userError_('Aloqa usulini tanlang', 'validation');
  var resList = Object.keys(CONTACT_RES).map(function (k) { return CONTACT_RES[k]; });
  if (resList.indexOf(d.natija) === -1) throw userError_('Natijani tanlang', 'validation');
  ensureSchema_();
  return withLock_(function () {
    var n = table_(SHEETS.NOMZODLAR), m = findRow_(n, id);
    if (!m) throw userError_('Nomzod topilmadi: ' + id, 'notfound');
    if (CONTACT_ALLOWED.indexOf(m['Holat']) === -1 && !(m['Holat'] === STATUS.STUDYING && !m['Boshlagan sana'])) {
      throw userError_('Bu nomzod holati "' + m['Holat'] + '", xabar berish kerak emas', 'validation');
    }
    var now = new Date();
    var sana = isoToDate_(d.sana);
    if (sana) sana.setHours(now.getHours(), now.getMinutes());
    var when = sana || now;
    var oqish = isoToDate_(d.oqish);
    var row = {};
    row['Sana'] = when; row['Nomzod ID'] = id; row["F.I.Sh."] = m["F.I.Sh."]; row['Usul'] = d.usul; row['Natija'] = d.natija;
    row["O'qish kuni"] = oqish || ''; row['Izoh'] = String(d.izoh || '').slice(0, 500); row['HR'] = user && user.name || '';
    appendObj_(sheet_(SHEETS.ALOQA), HEADERS.Aloqa, row);
    setCell_(n, m._row, 'Oxirgi aloqa', when);
    setCell_(n, m._row, 'Aloqa usuli', d.usul);
    setCell_(n, m._row, 'Aloqa natijasi', d.natija);
    if (oqish) setCell_(n, m._row, "O'qish kuni", oqish);
    var holat = m['Holat'];
    if (d.natija === CONTACT_RES.NO && holat !== STATUS.NOSHOW) {
      holat = STATUS.NOSHOW;
      setCell_(n, m._row, 'Kelmaslik sababi', 'Rad etdi' + (d.izoh ? ': ' + d.izoh : ''));
    } else if (d.natija === CONTACT_RES.YES && holat === STATUS.NOSHOW) {
      holat = m['Guruh'] ? STATUS.STUDYING : STATUS.NEW;   // qayta taklif qilindi
      setCell_(n, m._row, 'Kelmaslik sababi', '');
    }
    if (holat !== m['Holat']) setCell_(n, m._row, 'Holat', holat);
    setCell_(n, m._row, 'Yangilangan', now);
    return { ok: true, holat: holat };
  });
}

/** Qo'lda "O'qishga kelmadi" deb belgilash yoki bekor qilish (undo: true) */
function setNoShow_(d, user) {
  var id = String(d.id || '');
  ensureSchema_();
  return withLock_(function () {
    var n = table_(SHEETS.NOMZODLAR), m = findRow_(n, id);
    if (!m) throw userError_('Nomzod topilmadi: ' + id, 'notfound');
    var now = new Date(), holat;
    if (d.undo) {
      if (m['Holat'] !== STATUS.NOSHOW) throw userError_('Nomzod "O\'qishga kelmadi" holatida emas', 'validation');
      holat = m['Guruh'] ? STATUS.STUDYING : STATUS.NEW;
      setCell_(n, m._row, 'Kelmaslik sababi', '');
    } else {
      var inGroupNotStarted = m['Holat'] === STATUS.STUDYING && !m['Boshlagan sana'];
      if (m['Holat'] !== STATUS.NEW && !inGroupNotStarted) throw userError_('Faqat o\'qishni boshlamagan nomzodni belgilash mumkin', 'validation');
      holat = STATUS.NOSHOW;
      setCell_(n, m._row, 'Kelmaslik sababi', String(d.sabab || '').slice(0, 300));
    }
    setCell_(n, m._row, 'Holat', holat);
    setCell_(n, m._row, 'Yangilangan', now);
    return { ok: true, holat: holat };
  });
}


// ---------------------------------------------------------------------------
// 3-bosqich: o'quv boti (qo'llanmalarni nomzodlarga Telegram orqali yetkazish)
// ---------------------------------------------------------------------------
var WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbx-yrGtRqw4LaTrssUCaMdCKevyOScvHo3Oid5BO8mibSZMFc_Y9-MtlWgE3YW-7hQeIA/exec';
var LOG_ST = { SENT: 'Yuborildi', NOTLINKED: 'Ulanmagan', FAIL: 'Yetkazilmadi' };
var BTN = {
  MATS: "📚 Qo'llanmalar", PLAN: '📅 Jadval', ASK: '❓ Savol berish',
  SEND: '📤 Yuborish', REPORT: '📊 Hisobot', LIST: '📂 Materiallar'
};

/**
 * Bir marta ishga tushiring (OQUV_BOT_TOKEN qo'shilgan va yangi versiya deploy qilingandan keyin):
 * varaqlarni yaratadi va botni shu skriptga ulaydi.
 */
function oquvBotUlash() {
  if (!oquvToken_()) throw new Error("Avval Project Settings → Script Properties'ga OQUV_BOT_TOKEN qo'shing");
  setup();
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('OQUV_WEBHOOK_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, ''); props.setProperty('OQUV_WEBHOOK_KEY', key); }
  var me = tg_('getMe', {});
  if (!me.ok) throw new Error("OQUV_BOT_TOKEN noto'g'ri: " + (me.description || ''));
  if (me.result.id && String(props.getProperty('BOT_TOKEN') || '').indexOf(String(me.result.id) + ':') === 0) {
    throw new Error("OQUV_BOT_TOKEN HR botining tokeni bilan bir xil. O'quv boti uchun yangi bot oching");
  }
  var r = tg_('setWebhook', {
    url: WEB_APP_URL + '?bot=' + key, allowed_updates: ['message', 'callback_query'],
    drop_pending_updates: true, max_connections: 10
  });
  if (!r.ok) throw new Error("Botni ulab bo'lmadi: " + (r.description || ''));
  tg_('setMyCommands', { commands: [{ command: 'start', description: 'Boshlash' }] });
  tg_('setChatMenuButton', { menu_button: { type: 'web_app', text: 'Darslarim', web_app: { url: OQUV_APP_URL } } });
  tg_('setMyDescription', { description: "O'qish davridagi qo'llanmalar shu botga keladi. Boshlash uchun «Start» ni bosing." });
  props.setProperty('OQUV_BOT_USERNAME', me.result.username);
  var link = 'https://t.me/' + me.result.username;
  Logger.log('✅ Bot ulandi: @' + me.result.username + '\nNomzodlar uchun havola: ' + link);
  return link;
}

/** Muammo bo'lsa ishga tushiring: bot holatini jurnalga chiqaradi. */
function oquvBotHolati() {
  var info = tg_('getWebhookInfo', {});
  var r = info.result || {};
  Logger.log('Kutayotgan xabarlar: ' + r.pending_update_count + '\nOxirgi xato: ' + (r.last_error_message || "yo'q") +
    '\nUlangan manzil: ' + (r.url ? r.url.split('?')[0] : "yo'q"));
  return info;
}

function oquvToken_() { return PropertiesService.getScriptProperties().getProperty('OQUV_BOT_TOKEN'); }

function tgReq_(method, payload) {
  return {
    url: 'https://api.telegram.org/bot' + oquvToken_() + '/' + method,
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(payload)
  };
}

function tgParse_(res) {
  try { return JSON.parse(res.getContentText()); } catch (e) { return { ok: false }; }
}

function tg_(method, payload) {
  var q = tgReq_(method, payload);
  return tgParse_(UrlFetchApp.fetch(q.url, q));
}

/** Ko'p xabarni parallel yuboradi (Telegram cheklovi uchun 25 tadan, oralig'ida 1 soniya). */
function tgMany_(calls) {
  var out = [];
  for (var i = 0; i < calls.length; i += 25) {
    if (i > 0) Utilities.sleep(1000);
    var part = calls.slice(i, i + 25).map(function (c) { return tgReq_(c.method, c.payload); });
    UrlFetchApp.fetchAll(part).forEach(function (res) { out.push(tgParse_(res)); });
  }
  return out;
}

function send_(chat, text, markup) {
  var p = { chat_id: chat, text: text, disable_web_page_preview: true };
  if (markup) p.reply_markup = markup;
  return tg_('sendMessage', p);
}

function edit_(chat, mid, text, markup) {
  var p = { chat_id: chat, message_id: mid, text: text, disable_web_page_preview: true };
  if (markup) p.reply_markup = markup;
  return tg_('editMessageText', p);
}

function ikb_(rows) { return { inline_keyboard: rows }; }
function b_(text, data) { return { text: text, callback_data: data }; }

function traineeKb_() {
  return { keyboard: [[{ text: BTN.MATS }, { text: BTN.PLAN }], [{ text: BTN.ASK }]], resize_keyboard: true, is_persistent: true };
}
function hrKb_() {
  return { keyboard: [[{ text: BTN.SEND }, { text: BTN.REPORT }], [{ text: BTN.LIST }]], resize_keyboard: true, is_persistent: true };
}
function contactKb_() {
  return { keyboard: [[{ text: '📱 Telefon raqamni yuborish', request_contact: true }]], resize_keyboard: true, one_time_keyboard: true };
}

function hrIds_() { return String(getSetting_('HR Telegram ID') || '').split(/[,\s]+/).filter(String); }
function isHr_(id) { return hrIds_().indexOf(String(id)) > -1; }

/** O'quv boti orqali HR'ga xabar. Yetmasa (HR botga /start yozmagan bo'lsa) HR botidan yuboradi. */
function hrSend_(text, markup) {
  var ok = false;
  hrIds_().forEach(function (id) { if (send_(id, text, markup).ok) ok = true; });
  if (!ok) notify_(text + "\n\n(O'quv botiga /start yozing, shunda bunday xabarlar o'sha botdan keladi.)");
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// --- Kirish nuqtasi ---------------------------------------------------------
function oquvWebhook_(e) {
  var key = PropertiesService.getScriptProperties().getProperty('OQUV_WEBHOOK_KEY');
  if (!key || String(e.parameter.bot) !== key) return;
  var u;
  try { u = JSON.parse(e.postData.contents); } catch (err) { return; }
  // Telegram bir xabarni qayta yuborsa, ikkinchi marta ishlamaymiz
  var cache = CacheService.getScriptCache();
  var ck = 'upd_' + u.update_id;
  if (cache.get(ck)) return;
  cache.put(ck, '1', 21600);
  var chat = (u.message && u.message.chat && u.message.chat.id) || (u.callback_query && u.callback_query.from.id);
  try {
    if (u.callback_query) onCallback_(u.callback_query);
    else if (u.message && u.message.chat.type === 'private') onMessage_(u.message);
  } catch (err) {
    console.error(err);
    if (chat) send_(chat, (err && err.userMessage) || "Xatolik yuz berdi. Birozdan keyin qayta urinib ko'ring.");
  }
}

function onMessage_(m) {
  if (isHr_(m.from.id)) return onHrMessage_(m);
  return onTraineeMessage_(m);
}

function onCallback_(q) {
  var p = String(q.data || '').split('|');
  var chat = q.message && q.message.chat.id, mid = q.message && q.message.message_id;
  var answer = '';
  // Uzoq davom etadigan amallarda tugmani darhol "bo'shatamiz"
  var early = ['sy', 'sa', 'rm', 'd', 'ml'].indexOf(p[0]) > -1;
  if (early) tg_('answerCallbackQuery', { callback_query_id: q.id, text: '⏳' });
  try {
    answer = (isHr_(q.from.id) ? hrCallback_(p, chat, mid) : traineeCallback_(p, q.from, chat, mid)) || '';
  } finally {
    if (!early) tg_('answerCallbackQuery', { callback_query_id: q.id, text: answer });
  }
}

// --- Ma'lumotlar ------------------------------------------------------------
function ensureColumn_(name, header) {
  var sh = sheet_(name);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(header) === -1) sh.getRange(1, headers.join('') === '' ? 1 : lastCol + 1).setValue(header);
}

function traineeByTg_(tgId) {
  var found = null;
  table_(SHEETS.NOMZODLAR).rows.forEach(function (r) {
    if (r['Telegram ID'] !== '' && r['Telegram ID'] !== undefined && String(r['Telegram ID']) === String(tgId)) found = r;
  });
  return found;
}

function groupRow_(name) {
  return table_(SHEETS.GURUHLAR).rows.filter(function (g) { return String(g['Guruh']) === String(name); })[0] || null;
}

function dayNumber_(g) {
  var s = isoToDate_(isoDate_(g['Boshlanish sanasi'])), t = isoToDate_(isoDate_(new Date()));
  if (!s) return 0;
  return Math.round((t - s) / 86400000) + 1;
}

function shortDate_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd.MM') : String(v || '');
}

function materials_(day) {
  var sh = sheet_(SHEETS.QOLLANMALAR);
  if (!sh) return [];
  return table_(SHEETS.QOLLANMALAR).rows.filter(function (r) {
    return Number(r['Kun']) === Number(day) && r['Fayl ID'] !== '' && String(r['Faol']).trim().toLowerCase() !== "yo'q";
  }).sort(function (a, b) { return (Number(a['Tartib']) || 0) - (Number(b['Tartib']) || 0); });
}

function logRows_(group, day) {
  var sh = sheet_(SHEETS.YUBORISHLAR);
  if (!sh) return [];
  return table_(SHEETS.YUBORISHLAR).rows.filter(function (r) {
    return String(r['Guruh']) === String(group) && (day === undefined || Number(r['Kun']) === Number(day));
  });
}

function openedDays_(group) {
  var days = {};
  logRows_(group).forEach(function (r) { days[Number(r['Kun'])] = true; });
  return Object.keys(days).map(Number).sort(function (a, b) { return a - b; });
}

function members_(group) {
  return table_(SHEETS.NOMZODLAR).rows.filter(function (m) {
    return String(m['Guruh']) === String(group) && m['Holat'] !== STATUS.EXPELLED && m['Holat'] !== STATUS.REJECTED && m['Holat'] !== STATUS.NOSHOW;
  });
}

/** «Familiya Ism Sharif» → «Ism Familiya» (murojaat uchun). */
function nameSurname_(fio) {
  var p = String(fio || '').trim().split(/\s+/);
  return p.length > 1 ? p[1] + ' ' + p[0] : p[0];
}

function firstName_(fio) {
  var p = String(fio || '').trim().split(/\s+/);
  return p.length > 1 ? p[1] : p[0];
}

/** Yuborishlar varag'ida nomzod + kun qatorini yangilaydi yoki qo'shadi. */
function upsertLog_(log, group, day, m, holat) {
  var id = String(m['ID']);
  var r = log.rows.filter(function (x) { return String(x['Guruh']) === String(group) && Number(x['Kun']) === Number(day) && String(x['Nomzod ID']) === id; })[0];
  if (r) {
    if (r['Holat'] !== holat && !(holat !== LOG_ST.SENT && r['Holat'] === LOG_ST.SENT)) setCell_(log, r._row, 'Holat', holat);
    if (holat === LOG_ST.SENT) setCell_(log, r._row, 'Sana', new Date());
    return;
  }
  var row = {};
  row['Sana'] = new Date(); row['Guruh'] = group; row['Kun'] = Number(day); row['Nomzod ID'] = id;
  row["F.I.Sh."] = m["F.I.Sh."]; row['Holat'] = holat; row["O'qidi"] = '';
  appendObj_(log.sh, HEADERS.Yuborishlar, row);
  log.rows.push({ _row: log.sh.getLastRow(), 'Guruh': group, 'Kun': Number(day), 'Nomzod ID': id, 'Holat': holat });
}

// --- Materiallarni yetkazish ------------------------------------------------
var SEND_BY_TYPE = {
  'Video': ['sendVideo', 'video'], 'Rasm': ['sendPhoto', 'photo'],
  'Audio': ['sendAudio', 'audio'], 'Ovozli': ['sendVoice', 'voice']
};

function materialCall_(chat, mat) {
  var turi = String(mat['Turi']), name = String(mat['Nomi'] || '');
  if (turi === 'Havola') {
    return { method: 'sendMessage', payload: { chat_id: chat, text: '🔗 ' + String(mat['Fayl ID']), protect_content: true } };
  }
  var how = SEND_BY_TYPE[turi] || ['sendDocument', 'document'];
  var p = { chat_id: chat, caption: name ? '📄 ' + name : '', protect_content: true };
  p[how[1]] = String(mat['Fayl ID']);
  return { method: how[0], payload: p };
}

/**
 * Bir kunning materiallarini bir nechta nomzodga yuboradi. Har bir chatda tartib saqlanadi:
 * avval sarlavha, keyin materiallar, oxirida «O'qidim» tugmasi.
 * Natija: { chatId: true/false }
 */
function deliver_(chats, day, mats, opts) {
  opts = opts || {};
  var ok = { _bad: [] };
  chats.forEach(function (c) { ok[c] = true; });
  var rounds = [];
  rounds.push(function (c) {
    return { method: 'sendMessage', payload: { chat_id: c, text: '📘 ' + day + "-kun materiallari\nJami " + mats.length +
      " ta. O'qib bo'lgach, oxiridagi «✅ O'qidim» tugmasini bosing." + (opts.note ? '\n\n' + opts.note : '') } };
  });
  mats.forEach(function (mat) { rounds.push(function (c) { return materialCall_(c, mat); }); });
  if (!opts.noButton) {
    rounds.push(function (c) {
      return { method: 'sendMessage', payload: { chat_id: c, text: day + "-kun materiallarini o'qib bo'ldingizmi?",
        reply_markup: ikb_([[b_("✅ O'qidim", 'r|' + day)], [appBtn_()]]) } };
    });
  }
  rounds.forEach(function (make, i) {
    var live = chats.filter(function (c) { return ok[c]; });
    if (!live.length) return;
    var res = tgMany_(live.map(make));
    live.forEach(function (c, j) { if (!res[j].ok && i === 0) ok[c] = false; });
    var allFailed = res.length && res.every(function (r) { return !r.ok; });
    if (i > 0 && i <= mats.length && allFailed) ok._bad.push(String(mats[i - 1]['Nomi']) + ': ' + (res[0].description || 'xato'));
  });
  return ok;
}

// --- Nomzod tomoni ----------------------------------------------------------
function onTraineeMessage_(m) {
  var chat = m.chat.id, text = String(m.text || '').trim();
  if (m.contact) return registerTrainee_(m);
  var c = traineeByTg_(m.from.id);
  if (!c) {
    return send_(chat, "Assalomu alaykum! Bu o'quv boti: o'qish davridagi qo'llanmalar shu yerga keladi.\n\n" +
      "Boshlash uchun pastdagi «📱 Telefon raqamni yuborish» tugmasini bosing. Suhbatda qaysi raqamni yozgan bo'lsangiz, o'sha raqam kerak.", contactKb_());
  }
  if (c['Holat'] === STATUS.EXPELLED) {
    return send_(chat, "Siz o'qishdan chetlatilgansiz, shuning uchun qo'llanmalar yuborilmaydi. Savol bo'lsa, HR bilan bog'laning.", { remove_keyboard: true });
  }
  if (text.indexOf('/start') === 0) { welcome_(chat, c); return deliverPending_(c); }
  if (text === BTN.MATS) return showDays_(chat, c);
  if (text === BTN.PLAN) return showPlan_(chat, c);
  if (text === BTN.ASK) return send_(chat, 'Savolingizni bitta xabarda yozib yuboring. HR javob berganda shu yerga keladi.', traineeKb_());
  if (text) return forwardQuestion_(chat, c, text);
  return send_(chat, 'Pastdagi tugmalardan foydalaning.', traineeKb_());
}

function registerTrainee_(m) {
  var chat = m.chat.id, ct = m.contact;
  if (String(ct.user_id || '') !== String(m.from.id)) {
    return send_(chat, "Iltimos, boshqa odamning emas, o'zingizning raqamingizni pastdagi tugma orqali yuboring.", contactKb_());
  }
  var phone = normPhone_(ct.phone_number);
  var res = withLock_(function () {
    ensureColumn_(SHEETS.NOMZODLAR, 'Telegram ID');
    var n = table_(SHEETS.NOMZODLAR);
    var mine = n.rows.filter(function (r) { return normPhone_(r['Telefon']) === phone; });
    if (!mine.length) return { err: 'notfound' };
    var c = mine.filter(function (r) { return r['Guruh'] && r['Holat'] === STATUS.STUDYING; }).pop() ||
      mine.filter(function (r) { return r['Guruh']; }).pop() || mine[mine.length - 1];
    if (!c['Guruh']) return { err: 'nogroup', c: c };
    if (c['Holat'] === STATUS.EXPELLED) return { err: 'expelled', c: c };
    // Shu Telegram akkaunt boshqa qatorga bog'langan bo'lsa, o'sha bog'lanishni olib tashlaymiz
    n.rows.forEach(function (r) {
      if (r._row !== c._row && String(r['Telegram ID']) === String(m.from.id)) setCell_(n, r._row, 'Telegram ID', '');
    });
    setCell_(n, c._row, 'Telegram ID', String(m.from.id));
    c['Telegram ID'] = String(m.from.id);
    return { c: c };
  });
  if (res.err === 'notfound') {
    send_(chat, 'Bu raqam (' + fmtPhone_(phone) + ") nomzodlar ro'yxatida topilmadi.\n\nSuhbatda boshqa raqam yozgan bo'lsangiz, HR'ga ayting, u raqamni to'g'rilab beradi. Keyin qayta /start bosing.", { remove_keyboard: true });
    hrSend_('⚠️ Botga ro\'yxatda yo\'q raqam kirmoqchi bo\'ldi: ' + fmtPhone_(phone) +
      (m.from.first_name ? ' (' + [m.from.first_name, m.from.last_name].filter(String).join(' ') + ')' : '') +
      "\nAgar bu nomzod bo'lsa, Nomzodlar varag'ida uning telefonini shu raqamga o'zgartiring.");
    return;
  }
  if (res.err === 'nogroup') return send_(chat, "Siz hali o'quv guruhiga qo'shilmagansiz. Guruhga qo'shilgach, shu yerga /start yozing.", { remove_keyboard: true });
  if (res.err === 'expelled') return send_(chat, "Siz o'qishdan chetlatilgansiz, shuning uchun qo'llanmalar yuborilmaydi.", { remove_keyboard: true });
  welcome_(chat, res.c);
  deliverPending_(res.c);
}

function welcome_(chat, c) {
  var g = groupRow_(c['Guruh']) || {};
  var start = g['Boshlanish sanasi'], last = '';
  if (start instanceof Date) { last = new Date(start.getTime()); last.setDate(last.getDate() + TRAINING_DAYS - 1); }
  send_(chat, '✅ Xush kelibsiz, ' + firstName_(c["F.I.Sh."]) + '!\n\n' +
    'Guruh: ' + c['Guruh'] + '\n' +
    (start ? "O'qish: " + shortDate_(start) + ' – ' + shortDate_(last) + '\n' : '') +
    (g['Imtihon sanasi'] ? 'Imtihon: ' + fmtDate_(g['Imtihon sanasi']) + '\n' : '') +
    "\nQo'llanmalar shu botga keladi. Hammasini chiroyli ko'rinishda «Darslarim» ilovasida ochasiz.", traineeKb_());
  send_(chat, '📱 Darslar, materiallar va davomatingiz bitta ilovada:', ikb_([[appBtn_()]]));
}

/** Nomzod botga kech ulansa, HR allaqachon yuborgan kunlarni unga yetkazadi. */
function deliverPending_(c) {
  var id = String(c['ID']);
  // HR yuborgan, lekin bu nomzod hali olmagan kunlar (kech ulangan yoki guruhga keyin qo'shilgan)
  var all = logRows_(c['Guruh']), got = {}, days = {};
  all.forEach(function (r) { if (String(r['Nomzod ID']) === id && r['Holat'] === LOG_ST.SENT) got[Number(r['Kun'])] = true; });
  all.forEach(function (r) { var k = Number(r['Kun']); if (!got[k]) days[k] = true; });
  Object.keys(days).map(Number).sort(function (a, b) { return a - b; }).forEach(function (day) {
    var mats = materials_(day);
    if (!mats.length) return;
    var ok = deliver_([c['Telegram ID']], day, mats);
    if (ok[c['Telegram ID']]) withLock_(function () { upsertLog_(table_(SHEETS.YUBORISHLAR), c['Guruh'], day, c, LOG_ST.SENT); });
  });
}

function showDays_(chat, c) {
  var days = openedDays_(c['Guruh']);
  if (!days.length) return send_(chat, "Hozircha material yo'q. HR yuborganda shu yerga keladi.", traineeKb_());
  send_(chat, 'Qaysi kunning materiallarini ochamiz?', ikb_([days.map(function (d) { return b_(d + '-kun', 'd|' + d); })]));
}

function showPlan_(chat, c) {
  var g = groupRow_(c['Guruh']);
  if (!g) return send_(chat, "Guruh ma'lumoti topilmadi. HR bilan bog'laning.", traineeKb_());
  var n = dayNumber_(g), days = openedDays_(c['Guruh']);
  var lines = ['📅 Guruh ' + c['Guruh']];
  for (var i = 1; i <= TRAINING_DAYS; i++) {
    var d = new Date(g['Boshlanish sanasi'].getTime ? g['Boshlanish sanasi'].getTime() : 0); d.setDate(d.getDate() + i - 1);
    lines.push((i === n ? '👉 ' : (i < n ? '✔️ ' : '▫️ ')) + i + '-kun · ' + shortDate_(d) + (days.indexOf(i) > -1 ? ' · materiallar bor' : ''));
  }
  if (g['Imtihon sanasi']) lines.push('🎯 Imtihon · ' + fmtDate_(g['Imtihon sanasi']));
  send_(chat, lines.join('\n'), traineeKb_());
}

function ensureSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold').setBackground('#e2f4ea');
    sh.setFrozenRows(1);
  }
  return sh;
}

function forwardQuestion_(chat, c, text) {
  try {
    withLock_(function () {
      var row = {};
      row['Sana'] = new Date(); row['Nomzod ID'] = String(c['ID']); row["F.I.Sh."] = c["F.I.Sh."];
      row['Guruh'] = c['Guruh']; row['Savol'] = text;
      appendObj_(ensureSheet_(SHEETS.SAVOLLAR), HEADERS.Savollar, row);
    });
  } catch (e) { console.error(e); }
  hrSend_('❓ Savol\n👤 ' + c["F.I.Sh."] + '\n🆔 ' + c['ID'] + ' · Guruh ' + c['Guruh'] + '\n\n' + text +
    "\n\nJavob berish uchun shu xabarga «Ответить / Reply» qilib yozing.");
  send_(chat, "✅ Savolingiz HR'ga yuborildi. Javob shu yerga keladi.", traineeKb_());
}

function traineeCallback_(p, from, chat, mid) {
  var c = traineeByTg_(from.id);
  if (!c) return "Avval /start bosib, telefon raqamingizni yuboring";
  if (c['Holat'] === STATUS.EXPELLED) return "Siz o'qishdan chetlatilgansiz";
  var day = Number(p[1]);
  if (p[0] === 'd') {
    if (openedDays_(c['Guruh']).indexOf(day) === -1) return "Bu kun hali ochilmagan";
    var mats = materials_(day);
    if (!mats.length) return "Bu kun uchun material yo'q";
    var ok = deliver_([chat], day, mats);
    if (ok[chat]) withLock_(function () { upsertLog_(table_(SHEETS.YUBORISHLAR), c['Guruh'], day, c, LOG_ST.SENT); });
    return '';
  }
  if (p[0] === 'r') {
    markRead_(c, day);
    tg_('editMessageReplyMarkup', { chat_id: chat, message_id: mid, reply_markup: ikb_([[b_("✅ O'qildi", 'ok')]]) });
    return 'Rahmat! Belgilandi ✅';
  }
  if (p[0] === 'ok') return 'Allaqachon belgilangan';
  return '';
}

/** Nomzod + kun qatorini topadi, bo'lmasa «Yuborildi» bilan yaratadi. Qulf ichida chaqiriladi. */
function logRowFor_(log, c, day) {
  var find = function () {
    return log.rows.filter(function (x) { return String(x['Guruh']) === String(c['Guruh']) && Number(x['Kun']) === Number(day) && String(x['Nomzod ID']) === String(c['ID']); })[0];
  };
  if (!find()) upsertLog_(log, c['Guruh'], day, c, LOG_ST.SENT);
  return find();
}

function markRead_(c, day) {
  withLock_(function () {
    var log = table_(SHEETS.YUBORISHLAR);
    var r = logRowFor_(log, c, day);
    if (r && !r["O'qidi"]) setCell_(log, r._row, "O'qidi", new Date());
  });
}

// --- HR tomoni --------------------------------------------------------------
function fileOf_(m) {
  var cap = String(m.caption || '').trim();
  var base = function (fn) { return String(fn || '').replace(/\.[A-Za-z0-9]{1,5}$/, ''); };
  if (m.document) {
    var fn = String(m.document.file_name || ''), mime = String(m.document.mime_type || '');
    var t = /pdf/i.test(mime) || /\.pdf$/i.test(fn) ? 'PDF' :
      (/word|msword|officedocument\.wordprocessing/i.test(mime) || /\.docx?$/i.test(fn) ? 'Word' :
        (/^video\//i.test(mime) ? 'Video fayl' : 'Fayl'));
    return { turi: t, id: m.document.file_id, nomi: cap || base(fn) || t };
  }
  if (m.video) return { turi: 'Video', id: m.video.file_id, nomi: cap || base(m.video.file_name) || 'Video' };
  if (m.photo && m.photo.length) return { turi: 'Rasm', id: m.photo[m.photo.length - 1].file_id, nomi: cap || 'Rasm' };
  if (m.audio) return { turi: 'Audio', id: m.audio.file_id, nomi: cap || m.audio.title || base(m.audio.file_name) || 'Audio' };
  if (m.voice) return { turi: 'Ovozli', id: m.voice.file_id, nomi: cap || 'Ovozli xabar' };
  return null;
}

function onHrMessage_(m) {
  var chat = m.chat.id, text = String(m.text || '').trim();
  if (m.reply_to_message && text) {
    var ref = String(m.reply_to_message.text || '').match(/N-\d{4}-\d+/);
    if (ref) return hrAnswer_(chat, ref[0], text);
  }
  var f = fileOf_(m);
  if (!f && /https?:\/\//i.test(text)) f = { turi: 'Havola', id: text, nomi: text.replace(/https?:\/\/\S+/gi, '').trim() || 'Havola' };
  if (f) {
    CacheService.getScriptCache().put('pend_' + chat + '_' + m.message_id, JSON.stringify(f), 21600);
    var row = [];
    for (var d = 1; d <= TRAINING_DAYS; d++) row.push(b_(d + '-kun', 'a|' + m.message_id + '|' + d));
    return tg_('sendMessage', {
      chat_id: chat, reply_to_message_id: m.message_id, text: '«' + f.nomi + '» (' + f.turi + ")\nQaysi kunga qo'shay?",
      reply_markup: ikb_([row.slice(0, 3), row.slice(3), [b_('Bekor qilish', 'ax|' + m.message_id)]])
    });
  }
  if (text === BTN.SEND) return pickGroup_(chat, null, 'sg', '📤 Qaysi guruhga yuboramiz?');
  if (text === BTN.REPORT) return pickGroup_(chat, null, 'hg', '📊 Qaysi guruh hisoboti?');
  if (text === BTN.LIST) return listMaterials_(chat);
  var user = PropertiesService.getScriptProperties().getProperty('OQUV_BOT_USERNAME');
  send_(chat, "Assalomu alaykum! Siz HR sifatida kirdingiz.\n\n" +
    "➕ Material qo'shish: Word, PDF yoki videoni shu yerga yuboring yoki uzating, keyin kunni tanlang. Izoh (caption) yozsangiz, u material nomi bo'ladi. Havolani ham yuborsa bo'ladi.\n" +
    '📤 Yuborish: guruh va kunni tanlab, materiallarni nomzodlarga yuborasiz.\n' +
    "📊 Hisobot: kim o'qigani, kim botga ulanmagani.\n" +
    "📂 Materiallar: kunlar bo'yicha ro'yxat.\n" +
    "💬 Nomzod savoliga javob: savol xabariga «Reply» qilib yozing." +
    (user ? '\n\nNomzodlar uchun havola: https://t.me/' + user : ''), hrKb_());
}

function hrAnswer_(chat, id, text) {
  var c = table_(SHEETS.NOMZODLAR).rows.filter(function (r) { return String(r['ID']) === id; })[0];
  if (!c || !c['Telegram ID']) return send_(chat, '⚠️ ' + id + ' botga ulanmagan, javob yuborilmadi.', hrKb_());
  var r = send_(c['Telegram ID'], "💬 HR javobi:\n\n" + text, traineeKb_());
  if (r.ok) saveAnswer_(id, null, text);
  send_(chat, r.ok ? '✅ Javob yuborildi: ' + c["F.I.Sh."] : '⚠️ Yetkazib bo\'lmadi: ' + c["F.I.Sh."] + ' botni o\'chirgan bo\'lishi mumkin.', hrKb_());
}

function activeGroups_() {
  var n = table_(SHEETS.NOMZODLAR).rows;
  return table_(SHEETS.GURUHLAR).rows.map(function (g) {
    var name = String(g['Guruh']);
    var mine = n.filter(function (m) { return String(m['Guruh']) === name && m['Holat'] === STATUS.STUDYING; });
    return { name: name, g: g, count: mine.length };
  }).filter(function (x) { return x.count > 0 && String(x.g['Holat'] || '') !== 'Yopilgan'; }).reverse().slice(0, 8);
}

function pickGroup_(chat, mid, prefix, title) {
  var gs = activeGroups_();
  if (!gs.length) return send_(chat, "O'qishdagi guruh yo'q. Avval ilovada guruh yarating.", hrKb_());
  var rows = gs.map(function (x) {
    return [b_(x.name + ' · ' + shortDate_(x.g['Boshlanish sanasi']) + ' · ' + x.count + ' kishi', prefix + '|' + x.name)];
  });
  return mid ? edit_(chat, mid, title, ikb_(rows)) : send_(chat, title, ikb_(rows));
}

function pickDay_(chat, mid, group, prefix, title, onlyOpened) {
  var g = groupRow_(group);
  var today = g ? dayNumber_(g) : 0, opened = openedDays_(group), row = [];
  for (var d = 1; d <= TRAINING_DAYS; d++) {
    if (onlyOpened && opened.indexOf(d) === -1) continue;
    row.push(b_((d === today ? '👉 ' : '') + d + '-kun' + (!onlyOpened && opened.indexOf(d) > -1 ? ' ✓' : ''), prefix + '|' + group + '|' + d));
  }
  if (!row.length) return edit_(chat, mid, group + " guruhga hali material yuborilmagan.");
  edit_(chat, mid, title + (today >= 1 && today <= TRAINING_DAYS ? '\nBugun ' + today + '-kun.' : ''),
    ikb_([row.slice(0, 3), row.slice(3)].filter(function (r) { return r.length; })));
}

function hrCallback_(p, chat, mid) {
  var cache = CacheService.getScriptCache();
  switch (p[0]) {
    case 'a': {
      var key = 'pend_' + chat + '_' + p[1], raw = cache.get(key), day = Number(p[2]);
      if (!raw) { edit_(chat, mid, "Vaqti o'tib ketdi. Faylni qayta yuboring."); return ''; }
      var f = JSON.parse(raw);
      var res = withLock_(function () {
        var t = table_(SHEETS.QOLLANMALAR);
        var same = t.rows.filter(function (r) { return Number(r['Kun']) === day; });
        var order = same.reduce(function (mx, r) { return Math.max(mx, Number(r['Tartib']) || 0); }, 0) + 1;
        var row = {};
        row['Kun'] = day; row['Tartib'] = order; row['Nomi'] = f.nomi; row['Turi'] = f.turi;
        row['Fayl ID'] = f.id; row['Faol'] = 'Ha'; row["Qo'shilgan"] = new Date();
        appendObj_(t.sh, HEADERS["Qo'llanmalar"], row);
        return { order: order, total: materials_(day).length };
      });
      cache.remove(key);
      edit_(chat, mid, '✅ Saqlandi: ' + day + '-kun, ' + res.order + '-material\n«' + f.nomi + '»\n\n' + day + '-kunda jami ' + res.total + ' ta material.');
      return 'Saqlandi';
    }
    case 'ax':
      cache.remove('pend_' + chat + '_' + p[1]);
      edit_(chat, mid, 'Bekor qilindi.');
      return '';
    case 'x':
      edit_(chat, mid, 'Bekor qilindi.');
      return '';
    case 'ml':
      var mats = materials_(Number(p[1]));
      deliver_([chat], Number(p[1]), mats, { noButton: true, note: "(Bu sizga ko'rish uchun, nomzodlarga yuborilmadi.)" });
      return '';
    case 'sg': return pickDay_(chat, mid, p[1], 'sd', '📤 ' + p[1] + ' guruh. Qaysi kun materiallarini yuboramiz?', false), '';
    case 'hg': return pickDay_(chat, mid, p[1], 'hd', '📊 ' + p[1] + ' guruh. Qaysi kun?', true), '';
    case 'sd': return confirmSend_(chat, mid, p[1], Number(p[2])), '';
    case 'sy': return doSend_(chat, mid, p[1], Number(p[2]), false), '';
    case 'sa': return doSend_(chat, mid, p[1], Number(p[2]), true), '';
    case 'hd': return report_(chat, mid, p[1], Number(p[2])), '';
    case 'rm': return remind_(chat, p[1], Number(p[2]));
  }
  return '';
}

function listMaterials_(chat) {
  var lines = ['📂 Materiallar'], btns = [];
  for (var d = 1; d <= TRAINING_DAYS; d++) {
    var mats = materials_(d);
    lines.push('\n' + d + '-kun' + (mats.length ? ':' : ": yo'q"));
    mats.forEach(function (m, i) { lines.push('  ' + (i + 1) + '. ' + m['Nomi'] + ' (' + m['Turi'] + ')'); });
    if (mats.length) btns.push(b_('👁 ' + d + '-kun', 'ml|' + d));
  }
  lines.push("\nNomini, tartibini o'zgartirish yoki o'chirish: Sheets → «Qo'llanmalar» varag'i (Faol ustuniga «Yo'q» yozing).");
  send_(chat, lines.join('\n'), btns.length ? ikb_([btns.slice(0, 3), btns.slice(3)].filter(function (r) { return r.length; })) : hrKb_());
}

function confirmSend_(chat, mid, group, day) {
  var mats = materials_(day);
  if (!mats.length) return edit_(chat, mid, '⚠️ ' + day + "-kun uchun material yo'q. Avval faylni shu botga yuborib, " + day + '-kunni tanlang.');
  var ms = members_(group), log = logRows_(group, day);
  var sent = {};
  log.forEach(function (r) { if (r['Holat'] === LOG_ST.SENT) sent[String(r['Nomzod ID'])] = true; });
  var todo = ms.filter(function (m) { return !sent[String(m['ID'])]; });
  var linked = todo.filter(function (m) { return m['Telegram ID']; }).length;
  var lines = ['📤 ' + group + ' guruh, ' + day + '-kun', '', 'Materiallar (' + mats.length + '):'];
  mats.forEach(function (m, i) { lines.push((i + 1) + '. ' + m['Nomi'] + ' (' + m['Turi'] + ')'); });
  lines.push('');
  var rows = [];
  if (todo.length) {
    lines.push('Oladi: ' + linked + ' nomzod' + (todo.length - linked ? '\nBotga ulanmagan: ' + (todo.length - linked) + ' (ulanishi bilan avtomatik oladi)' : ''));
    if (ms.length - todo.length) lines.push('Oldin olgan: ' + (ms.length - todo.length) + ' (ularga qayta ketmaydi)');
    rows.push([b_('✅ Yuborish', 'sy|' + group + '|' + day), b_('Bekor qilish', 'x')]);
  } else {
    lines.push('Bu kun hamma ' + ms.length + " nomzodga yuborilgan. Yangi material qo'shgan bo'lsangiz, hammaga qayta yuborishingiz mumkin.");
    rows.push([b_('🔁 Hammaga qayta yuborish', 'sa|' + group + '|' + day), b_('Bekor qilish', 'x')]);
  }
  edit_(chat, mid, lines.join('\n'), ikb_(rows));
}

function doSend_(chat, mid, group, day, all) {
  if (!materials_(day).length) return edit_(chat, mid, '⚠️ ' + day + "-kun uchun material yo'q.");
  edit_(chat, mid, '⏳ ' + group + ' guruh, ' + day + '-kun yuborilmoqda…');
  var r = sendCore_(group, day, all);
  var good = r.good, notLinked = r.notLinked, failed = r.failed, ok = { _bad: r.bad };
  var lines = ['✅ ' + group + ' guruh, ' + day + '-kun yuborildi: ' + good.length + ' nomzod'];
  if (notLinked.length) lines.push('\n⚠️ Botga ulanmagan (' + notLinked.length + "), ulanishi bilan oladi:\n• " + notLinked.join('\n• '));
  if (failed.length) lines.push('\n❌ Yetkazilmadi (' + failed.length + "), botni o'chirgan bo'lishi mumkin:\n• " + failed.join('\n• '));
  if (ok._bad && ok._bad.length) lines.push("\n⚠️ Bu materiallar ketmadi, faylni botga qayta yuboring:\n• " + ok._bad.join('\n• '));
  edit_(chat, mid, lines.join('\n'));
}

/** Guruhga bir kunning materiallarini yuboradi. Natija: { good, notLinked, failed, bad } (ismlar) */
function sendCore_(group, day, all) {
  var mats = materials_(day);
  if (!mats.length) throw userError_(day + "-kun uchun material yo'q", 'validation');
  var ms = members_(group);
  var sent = {};
  logRows_(group, day).forEach(function (r) { if (r['Holat'] === LOG_ST.SENT) sent[String(r['Nomzod ID'])] = true; });
  var todo = ms.filter(function (m) { return all || !sent[String(m['ID'])]; });
  var linked = todo.filter(function (m) { return m['Telegram ID']; });
  var ok = linked.length ? deliver_(linked.map(function (m) { return String(m['Telegram ID']); }), day, mats) : {};
  var good = [], notLinked = [], failed = [];
  withLock_(function () {
    var log = table_(SHEETS.YUBORISHLAR);
    todo.forEach(function (m) {
      var st = !m['Telegram ID'] ? LOG_ST.NOTLINKED : (ok[String(m['Telegram ID'])] ? LOG_ST.SENT : LOG_ST.FAIL);
      upsertLog_(log, group, day, m, st);
      (st === LOG_ST.SENT ? good : (st === LOG_ST.NOTLINKED ? notLinked : failed)).push(m["F.I.Sh."]);
    });
  });
  return { good: good, notLinked: notLinked, failed: failed, bad: ok._bad || [] };
}

function report_(chat, mid, group, day) {
  var ms = members_(group), log = logRows_(group, day);
  var byId = {};
  log.forEach(function (r) { byId[String(r['Nomzod ID'])] = r; });
  var read = [], unread = [], notLinked = [];
  ms.forEach(function (m) {
    var r = byId[String(m['ID'])];
    if (!m['Telegram ID']) notLinked.push(m["F.I.Sh."]);
    else if (r && r["O'qidi"]) read.push(m["F.I.Sh."]);
    else unread.push(m["F.I.Sh."]);
  });
  var lines = ['📊 ' + group + ' guruh, ' + day + '-kun', "O'qidi: " + read.length + ' / ' + ms.length];
  if (unread.length) lines.push("\n⏳ O'qimagan (" + unread.length + '):\n• ' + unread.join('\n• '));
  if (notLinked.length) lines.push('\n⚠️ Botga ulanmagan (' + notLinked.length + '):\n• ' + notLinked.join('\n• '));
  if (read.length) lines.push("\n✅ O'qidi (" + read.length + '):\n• ' + read.join('\n• '));
  edit_(chat, mid, lines.join('\n'), unread.length ? ikb_([[b_("🔔 O'qimaganlarga eslatish", 'rm|' + group + '|' + day)]]) : null);
}

function remind_(chat, group, day) {
  send_(chat, '🔔 Eslatma yuborildi: ' + remindCore_(group, day) + ' nomzod.', hrKb_());
  return 'Yuborildi';
}

function remindCore_(group, day) {
  var log = logRows_(group, day), byId = {};
  log.forEach(function (r) { byId[String(r['Nomzod ID'])] = r; });
  var targets = members_(group).filter(function (m) {
    var r = byId[String(m['ID'])];
    return m['Telegram ID'] && !(r && r["O'qidi"]);
  });
  var res = tgMany_(targets.map(function (m) {
    return { method: 'sendMessage', payload: { chat_id: String(m['Telegram ID']),
      text: 'Assalomu alaykum va rohmatullohi va barakatuh, ' + nameSurname_(m["F.I.Sh."]) + '!\n\n' +
        '⏰ ' + day + "-kun materiallarini hali o'qib chiqmagansiz. Iltimos, ularni ko'rib chiqing va o'qib bo'lgach «✅ O'qidim» tugmasini bosing.\n\n" +
        "Materiallarni pastdagi «📱 Darslarim» tugmasi orqali bir bosishda ochasiz.",
      reply_markup: ikb_([[appBtn_()], [b_("✅ O'qidim", 'r|' + day)]]) } };
  }));
  return res.filter(function (r) { return r.ok; }).length;
}


// ---------------------------------------------------------------------------
// «Darslarim» — nomzodlar uchun mini ilova (o'quv botining menyu tugmasi)
// ---------------------------------------------------------------------------
var OQUV_APP_URL = 'https://abduqodirismoilov16-eng.github.io/nomzod-yoli/oquv/';
var MAX_FILE = 20 * 1024 * 1024; // Telegram bot faqat 20 MB gacha faylni yuklab bera oladi

function appBtn_() { return { text: '📱 Darslarim ilovasini ochish', web_app: { url: OQUV_APP_URL } }; }

function oquvApi_(req) {
  var token = oquvToken_();
  if (!token) throw userError_("O'quv boti hali sozlanmagan", 'config');
  var tg = verifyInitData_((req.auth || {}).initData || '', token, 24 * 3600);
  if (!tg) throw userError_("Ilovani o'quv botidagi «Darslarim» tugmasi orqali oching", 'auth');
  var bot = PropertiesService.getScriptProperties().getProperty('OQUV_BOT_USERNAME') || '';
  if (String(req.action).indexOf('oq.hr.') === 0) {
    if (!isHr_(tg.id)) throw userError_("Bu bo'lim faqat HR uchun", 'forbidden');
    return hrApi_(req.action, req.data || {}, tg, token);
  }
  var c = traineeByTg_(tg.id);
  if (!c) return { ok: true, linked: false, hr: isHr_(tg.id), bot: bot };
  if (c['Holat'] === STATUS.EXPELLED) return { ok: true, linked: true, expelled: true, bot: bot };
  var d = req.data || {};
  switch (req.action) {
    case 'oq.me': return oqMe_(c, bot);
    case 'oq.file': return oqFile_(c, d, token);
    case 'oq.chat': return oqChat_(c, d);
    case 'oq.read': {
      if (openedDays_(c['Guruh']).indexOf(Number(d.day)) === -1) throw userError_('Bu kun hali ochilmagan', 'validation');
      markRead_(c, Number(d.day));
      return { ok: true };
    }
    case 'oq.ask': {
      var text = String(d.text || '').trim();
      if (!text) throw userError_('Savolni yozing', 'validation');
      forwardQuestion_(c['Telegram ID'], c, text.slice(0, 1500));
      return { ok: true };
    }
  }
  throw userError_("Noma'lum amal", 'validation');
}

function oqMe_(c, bot) {
  var g = groupRow_(c['Guruh']) || {};
  var start = g['Boshlanish sanasi'] instanceof Date ? g['Boshlanish sanasi'] : null;
  var opened = openedDays_(c['Guruh']);
  var id = String(c['ID']);
  var log = {};
  logRows_(c['Guruh']).forEach(function (r) { if (String(r['Nomzod ID']) === id) log[Number(r['Kun'])] = r; });
  var att = {};
  table_(SHEETS.DAVOMAT).rows.forEach(function (a) {
    if (String(a['Guruh']) === String(c['Guruh']) && String(a['Nomzod ID']) === id) att[Number(a['Kun'])] = a['Holat'];
  });
  var days = [];
  for (var d = 1; d <= TRAINING_DAYS; d++) {
    var date = start ? new Date(start.getTime()) : null;
    if (date) date.setDate(date.getDate() + d - 1);
    var isOpen = opened.indexOf(d) > -1;
    var seen = String((log[d] && log[d]['Ochilgan']) || '').split(/\s+/);
    days.push({
      day: d, date: date ? isoDate_(date) : '', open: isOpen,
      read: !!(log[d] && log[d]["O'qidi"]), att: att[d] || '',
      mats: isOpen ? materials_(d).map(function (m) {
        var t = Number(m['Tartib']);
        return { t: t, nomi: String(m['Nomi'] || ''), turi: String(m['Turi'] || ''), seen: seen.indexOf('#' + t) > -1 };
      }) : []
    });
  }
  return {
    ok: true, linked: true, bot: bot,
    me: {
      name: c["F.I.Sh."], first: firstName_(c["F.I.Sh."]), group: c['Guruh'],
      start: start ? isoDate_(start) : '', exam: isoDate_(g['Imtihon sanasi']), today: start ? dayNumber_(g) : 0
    },
    days: days
  };
}

function oqMaterial_(c, d) {
  var day = Number(d.day), t = Number(d.t);
  if (c && openedDays_(c['Guruh']).indexOf(day) === -1) throw userError_('Bu kun hali ochilmagan', 'validation');
  var m = materials_(day).filter(function (x) { return Number(x['Tartib']) === t; })[0];
  if (!m) throw userError_("Material topilmadi. Ro'yxatni yangilang", 'validation');
  return { day: day, t: t, m: m };
}

function markOpened_(c, day, t) {
  withLock_(function () {
    ensureColumn_(SHEETS.YUBORISHLAR, 'Ochilgan');
    var log = table_(SHEETS.YUBORISHLAR);
    var r = logRowFor_(log, c, day);
    if (!r) return;
    var v = String(r['Ochilgan'] || '').trim();
    var tags = v ? v.split(/\s+/) : [];
    if (tags.indexOf('#' + t) === -1) { tags.push('#' + t); setCell_(log, r._row, 'Ochilgan', tags.join(' ')); }
  });
}

/** Material qanday ko'rsatiladi: pdf, docx, video, image, audio, link yoki chat (ilovada ochib bo'lmaydi). */
function viewKind_(turi, path) {
  var ext = (String(path || '').match(/\.([a-z0-9]+)$/i) || ['', ''])[1].toLowerCase();
  if (turi === 'Havola') return 'link';
  if (ext === 'pdf') return 'pdf';
  if (ext === 'docx') return 'docx';
  if (['mp4', 'm4v', 'mov', 'webm'].indexOf(ext) > -1 || turi === 'Video') return 'video';
  if (['jpg', 'jpeg', 'png', 'webp', 'gif'].indexOf(ext) > -1 || turi === 'Rasm') return 'image';
  if (['mp3', 'm4a', 'aac', 'wav'].indexOf(ext) > -1 || turi === 'Audio') return 'audio';
  return 'chat';
}

var VIEW_MIME = { pdf: 'application/pdf', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', image: 'image/jpeg', audio: 'audio/mpeg' };

function oqFile_(c, d, token) {
  var x = oqMaterial_(c, d), m = x.m;
  if (c) markOpened_(c, x.day, x.t);
  var base = { ok: true, name: String(m['Nomi'] || ''), turi: String(m['Turi'] || '') };
  if (m['Turi'] === 'Havola') {
    var txt = String(m['Fayl ID']);
    base.kind = 'link'; base.url = (txt.match(/https?:\/\/\S+/) || [''])[0]; base.text = txt;
    return base;
  }
  var info = tg_('getFile', { file_id: String(m['Fayl ID']) });
  var path = info.ok && info.result ? info.result.file_path : '';
  var kind = viewKind_(m['Turi'], path);
  var size = info.ok && info.result ? Number(info.result.file_size || 0) : 0;
  if (!info.ok || !path || size > MAX_FILE || kind === 'chat') {
    base.kind = 'chat';
    base.reason = (!info.ok || size > MAX_FILE) ? 'big' : 'format';
    return base;
  }
  var res = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + token + '/' + path, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) { base.kind = 'chat'; base.reason = 'big'; return base; }
  var ext = (path.match(/\.([a-z0-9]+)$/i) || ['', ''])[1].toLowerCase();
  base.kind = kind;
  base.mime = kind === 'video' ? (ext === 'webm' ? 'video/webm' : (ext === 'mov' ? 'video/quicktime' : 'video/mp4')) :
    (kind === 'image' && ext === 'png' ? 'image/png' : VIEW_MIME[kind]);
  base.data = Utilities.base64Encode(res.getBlob().getBytes());
  return base;
}

function oqChat_(c, d) {
  var x = oqMaterial_(c, d);
  var r = tgMany_([materialCall_(String(c['Telegram ID']), x.m)])[0];
  if (!r || !r.ok) throw userError_("Chatga yuborib bo'lmadi. Botni qayta ishga tushiring (/start)", 'tg');
  markOpened_(c, x.day, x.t);
  return { ok: true };
}

// --- «Darslarim» ichidagi HR paneli -----------------------------------------
function saveAnswer_(id, row, text) {
  withLock_(function () {
    ensureSheet_(SHEETS.SAVOLLAR);
    var t = table_(SHEETS.SAVOLLAR);
    var r = row ? t.rows.filter(function (x) { return x._row === Number(row); })[0]
      : t.rows.filter(function (x) { return String(x['Nomzod ID']) === String(id) && !x['Javob']; }).pop();
    if (!r) return;
    setCell_(t, r._row, 'Javob', text);
    setCell_(t, r._row, 'Javob sanasi', new Date());
  });
}

function hrApi_(action, d, tg, token) {
  switch (action) {
    case 'oq.hr.home': return hrHome_(tg);
    case 'oq.hr.mats': return { ok: true, days: hrMats_() };
    case 'oq.hr.matSave': return hrMatSave_(d);
    case 'oq.hr.matMove': return hrMatMove_(d);
    case 'oq.hr.upload': return hrUpload_(d, tg, token);
    case 'oq.hr.preview': return oqFile_(null, d, token);
    case 'oq.hr.sendInfo': return hrSendInfo_(d);
    case 'oq.hr.send': return { ok: true, result: sendCore_(String(d.group), Number(d.day), !!d.all) };
    case 'oq.hr.report': return hrReport_(d);
    case 'oq.hr.remind': return { ok: true, count: remindCore_(String(d.group), Number(d.day)) };
    case 'oq.hr.questions': return hrQuestions_();
    case 'oq.hr.answer': return hrAnswerApi_(d);
  }
  throw userError_("Noma'lum amal", 'validation');
}

function hrHome_(tg) {
  var n = table_(SHEETS.NOMZODLAR).rows;
  var groups = activeGroups_().map(function (x) {
    var mine = n.filter(function (m) { return String(m['Guruh']) === x.name && m['Holat'] === STATUS.STUDYING; });
    return {
      name: x.name, start: isoDate_(x.g['Boshlanish sanasi']), exam: isoDate_(x.g['Imtihon sanasi']),
      count: x.count, linked: mine.filter(function (m) { return m['Telegram ID']; }).length,
      today: dayNumber_(x.g), opened: openedDays_(x.name)
    };
  });
  var q = sheet_(SHEETS.SAVOLLAR) ? table_(SHEETS.SAVOLLAR).rows.filter(function (r) { return !r['Javob']; }).length : 0;
  var counts = [];
  for (var d = 1; d <= TRAINING_DAYS; d++) counts.push(materials_(d).length);
  return {
    ok: true, hrPanel: true, name: [tg.first_name, tg.last_name].filter(String).join(' '),
    bot: PropertiesService.getScriptProperties().getProperty('OQUV_BOT_USERNAME') || '',
    groups: groups, unanswered: q, matCounts: counts
  };
}

function hrMats_() {
  var rows = sheet_(SHEETS.QOLLANMALAR) ? table_(SHEETS.QOLLANMALAR).rows : [];
  var days = [];
  for (var d = 1; d <= TRAINING_DAYS; d++) {
    days.push({
      day: d,
      mats: rows.filter(function (r) { return Number(r['Kun']) === d && r['Fayl ID'] !== ''; })
        .sort(function (a, b) { return (Number(a['Tartib']) || 0) - (Number(b['Tartib']) || 0); })
        .map(function (r) {
          return { row: r._row, t: Number(r['Tartib']), nomi: String(r['Nomi'] || ''), turi: String(r['Turi'] || ''),
            faol: String(r['Faol']).trim().toLowerCase() !== "yo'q" };
        })
    });
  }
  return days;
}

function matRowCheck_(t, d) {
  var r = t.rows.filter(function (x) { return x._row === Number(d.row); })[0];
  if (!r || (d.t !== undefined && Number(r['Tartib']) !== Number(d.t))) throw userError_("Ro'yxat o'zgargan. Sahifani yangilang", 'stale');
  return r;
}

function hrMatSave_(d) {
  withLock_(function () {
    var t = table_(SHEETS.QOLLANMALAR), r = matRowCheck_(t, d);
    if (d.nomi !== undefined) {
      var nm = String(d.nomi).trim();
      if (!nm) throw userError_('Nomini yozing', 'validation');
      setCell_(t, r._row, 'Nomi', nm.slice(0, 120));
    }
    if (d.faol !== undefined) setCell_(t, r._row, 'Faol', d.faol ? 'Ha' : "Yo'q");
    if (d.kun !== undefined && Number(d.kun) !== Number(r['Kun'])) {
      var k = Number(d.kun);
      if (!(k >= 1 && k <= TRAINING_DAYS)) throw userError_("Kun noto'g'ri", 'validation');
      var max = t.rows.filter(function (x) { return Number(x['Kun']) === k; }).reduce(function (m, x) { return Math.max(m, Number(x['Tartib']) || 0); }, 0);
      setCell_(t, r._row, 'Kun', k);
      setCell_(t, r._row, 'Tartib', max + 1);
    }
  });
  return { ok: true, days: hrMats_() };
}

function hrMatMove_(d) {
  withLock_(function () {
    var t = table_(SHEETS.QOLLANMALAR), r = matRowCheck_(t, d);
    var same = t.rows.filter(function (x) { return Number(x['Kun']) === Number(r['Kun']) && x['Fayl ID'] !== ''; })
      .sort(function (a, b) { return (Number(a['Tartib']) || 0) - (Number(b['Tartib']) || 0); });
    var i = same.indexOf(r), j = i + (Number(d.dir) < 0 ? -1 : 1);
    if (j < 0 || j >= same.length) return;
    var tmp = same[i]; same[i] = same[j]; same[j] = tmp;
    same.forEach(function (x, k) { if (Number(x['Tartib']) !== k + 1) setCell_(t, x._row, 'Tartib', k + 1); });
  });
  return { ok: true, days: hrMats_() };
}

function hrUpload_(d, tg, token) {
  var day = Number(d.kun);
  if (!(day >= 1 && day <= TRAINING_DAYS)) throw userError_('Kunni tanlang', 'validation');
  var fname = String(d.name || 'fayl'), mime = String(d.mime || 'application/octet-stream');
  if (!d.data) throw userError_('Fayl tanlanmagan', 'validation');
  var ext = (fname.match(/\.([a-z0-9]+)$/i) || ['', ''])[1].toLowerCase();
  var isVideo = /^video\//.test(mime) || ['mp4', 'mov', 'm4v', 'webm'].indexOf(ext) > -1;
  var turi = ext === 'pdf' ? 'PDF' : (ext === 'docx' || ext === 'doc' ? 'Word' : (isVideo ? 'Video' :
    (['jpg', 'jpeg', 'png', 'webp'].indexOf(ext) > -1 ? 'Rasm (fayl)' : 'Fayl')));
  var blob = Utilities.newBlob(Utilities.base64Decode(String(d.data)), mime, fname);
  var field = isVideo ? 'video' : 'document';
  var payload = { chat_id: String(tg.id), disable_notification: 'true', caption: "📎 Darslarim: " + day + "-kunga qo'shildi" };
  payload[field] = blob;
  if (isVideo) payload.supports_streaming = 'true';
  var res = tgParse_(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + (isVideo ? 'sendVideo' : 'sendDocument'),
    { method: 'post', payload: payload, muteHttpExceptions: true }));
  if (!res.ok) throw userError_("Faylni Telegram'ga yuklab bo'lmadi: " + (res.description || '') + ". Botga to'g'ridan yuborib ko'ring", 'tg');
  var msg = res.result, fid = (msg.video && msg.video.file_id) || (msg.document && msg.document.file_id) || (msg.animation && msg.animation.file_id);
  if (!fid) throw userError_("Telegram fayl ID bermadi. Botga to'g'ridan yuborib ko'ring", 'tg');
  if (!msg.video && isVideo) turi = 'Video fayl';
  tg_('deleteMessage', { chat_id: String(tg.id), message_id: msg.message_id });
  var nomi = String(d.nomi || '').trim() || fname.replace(/\.[A-Za-z0-9]{1,5}$/, '');
  withLock_(function () {
    ensureSheet_(SHEETS.QOLLANMALAR);
    var t = table_(SHEETS.QOLLANMALAR);
    var max = t.rows.filter(function (x) { return Number(x['Kun']) === day; }).reduce(function (m, x) { return Math.max(m, Number(x['Tartib']) || 0); }, 0);
    var row = {};
    row['Kun'] = day; row['Tartib'] = max + 1; row['Nomi'] = nomi.slice(0, 120); row['Turi'] = turi;
    row['Fayl ID'] = fid; row['Faol'] = 'Ha'; row["Qo'shilgan"] = new Date();
    appendObj_(t.sh, HEADERS["Qo'llanmalar"], row);
  });
  return { ok: true, days: hrMats_() };
}

function hrSendInfo_(d) {
  var group = String(d.group), day = Number(d.day);
  var mats = materials_(day), ms = members_(group), sent = {};
  logRows_(group, day).forEach(function (r) { if (r['Holat'] === LOG_ST.SENT) sent[String(r['Nomzod ID'])] = true; });
  var todo = ms.filter(function (m) { return !sent[String(m['ID'])]; });
  return {
    ok: true, mats: mats.map(function (m) { return { nomi: String(m['Nomi']), turi: String(m['Turi']) }; }),
    total: ms.length, todo: todo.length, linked: todo.filter(function (m) { return m['Telegram ID']; }).length,
    notLinked: todo.filter(function (m) { return !m['Telegram ID']; }).map(function (m) { return m["F.I.Sh."]; })
  };
}

function hrReport_(d) {
  var group = String(d.group);
  var ms = members_(group), log = logRows_(group), opened = openedDays_(group);
  var byKey = {};
  log.forEach(function (r) { byKey[String(r['Nomzod ID']) + '|' + Number(r['Kun'])] = r; });
  var matN = {};
  opened.forEach(function (k) { matN[k] = materials_(k).length; });
  return {
    ok: true, group: group, opened: opened, matN: matN,
    people: ms.map(function (m) {
      var id = String(m['ID']), days = {};
      opened.forEach(function (k) {
        var r = byKey[id + '|' + k];
        var seen = r ? String(r['Ochilgan'] || '').split(/\s+/).filter(String).length : 0;
        days[k] = { st: r ? String(r['Holat']) : '', read: !!(r && r["O'qidi"]), seen: seen };
      });
      return { name: m["F.I.Sh."], id: id, linked: !!m['Telegram ID'], days: days };
    })
  };
}

function hrQuestions_() {
  if (!sheet_(SHEETS.SAVOLLAR)) return { ok: true, items: [] };
  var tz = Session.getScriptTimeZone();
  return {
    ok: true, items: table_(SHEETS.SAVOLLAR).rows.slice(-60).reverse().map(function (r) {
      return {
        row: r._row, date: r['Sana'] instanceof Date ? Utilities.formatDate(r['Sana'], tz, 'dd.MM HH:mm') : String(r['Sana'] || ''),
        name: String(r["F.I.Sh."] || ''), group: String(r['Guruh'] || ''), savol: String(r['Savol'] || ''), javob: String(r['Javob'] || '')
      };
    })
  };
}

function hrAnswerApi_(d) {
  var text = String(d.text || '').trim();
  if (!text) throw userError_('Javobni yozing', 'validation');
  var t = table_(SHEETS.SAVOLLAR);
  var r = t.rows.filter(function (x) { return x._row === Number(d.row); })[0];
  if (!r) throw userError_("Savol topilmadi. Sahifani yangilang", 'stale');
  var c = table_(SHEETS.NOMZODLAR).rows.filter(function (x) { return String(x['ID']) === String(r['Nomzod ID']); })[0];
  if (!c || !c['Telegram ID']) throw userError_("Nomzod botga ulanmagan, javob yetkazilmaydi", 'validation');
  var res = send_(c['Telegram ID'], "💬 HR javobi:\n\n❓ " + String(r['Savol']).slice(0, 300) + "\n\n" + text, traineeKb_());
  if (!res.ok) throw userError_("Yetkazib bo'lmadi: nomzod botni o'chirgan bo'lishi mumkin", 'tg');
  saveAnswer_(r['Nomzod ID'], r._row, text);
  return hrQuestions_();
}
