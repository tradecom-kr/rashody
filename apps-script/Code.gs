/**
 * Расходы команды — сервер на Google Apps Script.
 * Данные хранятся в этой таблице на листах: Сотрудники, Расходы, Выдачи, Настройки.
 * Приложение (сайт) обращается сюда через «Веб-приложение» (Развернуть → Новое развертывание).
 *
 * Все проверки паролей — здесь, на сервере. Пароли в таблице видите только вы.
 */

var TZ = 'Asia/Seoul';
var SALT = 'rashody-komandy-v1';
var SH = { staff: 'Сотрудники', exp: 'Расходы', pay: 'Выдачи', set: 'Настройки' };
var HEAD = {
  staff: ['ID', 'Имя', 'Бюджет', 'Пароль', 'Порядок', 'Скрыт', 'Хеш пароля'],
  exp:   ['ID', 'Дата', 'ID сотрудника', 'Сотрудник', 'Категория', 'Сумма', 'Комментарий', 'Внесено', 'Место'],
  pay:   ['ID', 'Дата', 'ID сотрудника', 'Сотрудник', 'Сумма', 'Кто отметил', 'Внесено']
};
var DEFAULT_CATS = 'Бензин, Еда, Транспорт, Прочее';

/* ---------- вход ---------- */

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return handle_(body);
}

function handle_(req) {
  var out;
  try {
    out = route_(req);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function route_(req) {
  var a = req.action || 'data';
  if (a === 'data') return { ok: true, data: readAll_() };
  if (a === 'login') return login_(req);
  if (a === 'roleLogin') return { ok: true, role: role_(req.rolePass) || fail_('Неверный пароль.') };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (a === 'add') return add_(req);
    if (a === 'del') return del_(req);
    if (a === 'pay') return pay_(req);
    if (a === 'unpay') return unpay_(req);
    fail_('Неизвестное действие.');
  } finally {
    lock.releaseLock();
  }
}

function fail_(msg) { throw new Error(msg); }

/* ---------- таблица ---------- */

function sheet_(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH[key]);
  if (!sh) {
    sh = ss.insertSheet(SH[key]);
    if (HEAD[key]) {
      sh.getRange(1, 1, 1, HEAD[key].length).setValues([HEAD[key]]).setFontWeight('bold').setBackground('#E7E6E6');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// добавить недостающие столбцы в шапку (например «Место» в старой таблице)
function ensureHead_(key) {
  var sh = sheet_(key), want = HEAD[key]; if (!want) return sh;
  var have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  if (have.length < want.length) sh.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold').setBackground('#E7E6E6');
  return sh;
}

function rows_(key) {
  var sh = sheet_(key);
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var h = v[0];
  return v.slice(1).map(function (r, i) {
    var o = { _row: i + 2 };
    h.forEach(function (name, c) { o[name] = r[c]; });
    return o;
  }).filter(function (o) { return o['ID'] !== '' && o['ID'] !== null; });
}

function settings_() {
  var sh = sheet_('set');
  var v = sh.getDataRange().getValues();
  var s = {};
  v.forEach(function (r) { if (r[0]) s[String(r[0]).trim()] = r[1]; });
  return s;
}

function iso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var t = String(v || '').trim().replace(/\./g, '-');
  return /^\d{4}-\d{2}-\d{2}$/.test(t) ? t : '';
}

function dateCell_(iso) {
  return Utilities.parseDate(iso + ' 12:00', TZ, 'yyyy-MM-dd HH:mm'); // полдень по Сеулу — дата не съедет при любом часовом поясе
}

function num_(v) { var n = Number(String(v).replace(/[^\d.-]/g, '')); return isFinite(n) ? n : 0; }

function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }

function hash_(id, pin) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + ':' + id + ':' + pin, Utilities.Charset.UTF_8);
  return b.map(function (x) { var h = (x & 0xff).toString(16); return h.length === 1 ? '0' + h : h; }).join('');
}

/* ---------- чтение ---------- */

function staffAll_() {
  return rows_('staff').map(function (r) {
    return {
      id: String(r['ID']),
      name: String(r['Имя'] || 'Сотрудник'),
      budget: num_(r['Бюджет']),
      order: num_(r['Порядок']),
      hidden: r['Скрыт'] === true || String(r['Скрыт']).toLowerCase() === 'да' || String(r['Скрыт']).toLowerCase() === 'true',
      pin: String(r['Пароль'] || '').trim(),
      pinHash: String(r['Хеш пароля'] || '').trim()
    };
  }).sort(function (a, b) { return a.order - b.order; });
}

function readAll_() {
  var st = settings_();
  var cats = String(st['Категории'] || DEFAULT_CATS).split(',').map(function (x) { return x.trim(); }).filter(String);
  return {
    staff: staffAll_().map(function (s) {
      return { id: s.id, name: s.name, budget: s.budget, order: s.order, hidden: s.hidden, hasPin: !!(s.pin || s.pinHash) };
    }),
    cats: cats,
    expenses: rows_('exp').map(function (r) {
      return { id: String(r['ID']), date: iso_(r['Дата']), staff: String(r['ID сотрудника']), cat: String(r['Категория'] || ''),
               amount: num_(r['Сумма']), note: String(r['Комментарий'] || ''), place: String(r['Место'] || ''), created: r['Внесено'] instanceof Date ? r['Внесено'].getTime() : 0 };
    }).filter(function (e) { return e.amount > 0; }),
    payments: rows_('pay').map(function (r) {
      return { id: String(r['ID']), date: iso_(r['Дата']), staff: String(r['ID сотрудника']), amount: num_(r['Сумма']),
               by: String(r['Кто отметил'] || ''), created: r['Внесено'] instanceof Date ? r['Внесено'].getTime() : 0 };
    }).filter(function (p) { return p.amount !== 0; }),
    serverTime: Date.now()
  };
}

/* ---------- проверки ---------- */

function staffOk_(id, pin) {
  var s = staffAll_().filter(function (x) { return x.id === String(id) && !x.hidden; })[0];
  if (!s) fail_('Сотрудник не найден.');
  if (!s.pin && !s.pinHash) return s;                      // пароль не задан
  pin = String(pin || '');
  if (s.pin && pin === s.pin) return s;
  if (s.pinHash && hash_(s.id, pin) === s.pinHash) return s;
  fail_('Неверный пароль сотрудника.');
}

function role_(pass) {
  pass = String(pass || '');
  if (!pass) return '';
  var st = settings_();
  // Бухгалтер — высшая роль (полные права). Старая строка «Пароль руководителя» тоже работает, если заполнена.
  if (String(st['Пароль бухгалтера'] || '') && pass === String(st['Пароль бухгалтера'])) return 'owner';
  if (String(st['Пароль руководителя'] || '') && pass === String(st['Пароль руководителя'])) return 'owner';
  return '';
}

function login_(req) {
  var s = staffOk_(req.staff, req.pin);
  return { ok: true, staff: s.id };
}

/* ---------- действия ---------- */

function add_(req) {
  var s;
  var role = role_(req.rolePass);
  if (role === 'owner') {
    s = staffAll_().filter(function (x) { return x.id === String(req.staff); })[0] || fail_('Сотрудник не найден.');
  } else {
    s = staffOk_(req.staff, req.pin);
  }
  var amount = Math.round(num_(req.amount));
  if (!(amount > 0)) fail_('Введите сумму.');
  if (amount > 100000000) fail_('Сумма слишком большая.');
  var date = iso_(req.date) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (date > today) fail_('Дата не может быть в будущем.');
  var cat = String(req.cat || '').slice(0, 40);
  var note = String(req.note || '').slice(0, 120);
  var place = String(req.place || '').slice(0, 80);
  // ID приходит с телефона: если связь оборвалась и запрос повторили — дубль не создаётся
  var cid = String(req.clientId || '');
  var id = /^[A-Za-z0-9_-]{8,40}$/.test(cid) ? cid : newId_();
  if (cid && rows_('exp').some(function (r) { return String(r['ID']) === id; })) return { ok: true, id: id, dup: true };
  var sh = ensureHead_('exp');
  sh.appendRow([id, dateCell_(date), s.id, s.name, cat, amount, note, new Date(), place]);
  var r = sh.getLastRow();
  sh.getRange(r, 2).setNumberFormat('yyyy.mm.dd');
  sh.getRange(r, 6).setNumberFormat('#,##0');
  sh.getRange(r, 8).setNumberFormat('yyyy.mm.dd hh:mm');
  return { ok: true, id: id };
}

function del_(req) {
  var row = rows_('exp').filter(function (r) { return String(r['ID']) === String(req.id); })[0];
  if (!row) fail_('Запись не найдена.');
  if (role_(req.rolePass) !== 'owner') {
    var s = staffOk_(req.staff, req.pin);
    if (String(row['ID сотрудника']) !== s.id) fail_('Удалять можно только свои расходы.');
  }
  sheet_('exp').deleteRow(row._row);
  return { ok: true };
}

function pay_(req) {
  var role = role_(req.rolePass);
  if (!role) fail_('Выдавать может только бухгалтер.');
  var s = staffAll_().filter(function (x) { return x.id === String(req.staff); })[0] || fail_('Сотрудник не найден.');
  var amount = Math.round(num_(req.amount));
  if (!(amount > 0)) fail_('Нечего выдавать.');
  var id = newId_();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var sh = sheet_('pay');
  sh.appendRow([id, dateCell_(today), s.id, s.name, amount, 'Бухгалтер', new Date()]);
  var r = sh.getLastRow();
  sh.getRange(r, 2).setNumberFormat('yyyy.mm.dd');
  sh.getRange(r, 5).setNumberFormat('#,##0');
  sh.getRange(r, 7).setNumberFormat('yyyy.mm.dd hh:mm');
  return { ok: true, id: id };
}

function unpay_(req) {
  var role = role_(req.rolePass);
  if (!role) fail_('Отменять выдачу может только бухгалтер.');
  var row = rows_('pay').filter(function (r) { return String(r['ID']) === String(req.id); })[0];
  if (!row) fail_('Выдача не найдена.');
  if (role === 'acc' && String(row['Кто отметил']) !== 'Бухгалтер') fail_('Бухгалтер может отменить только свою выдачу.');
  sheet_('pay').deleteRow(row._row);
  return { ok: true };
}

/* ---------- первый запуск: создать листы, если их нет ---------- */

function setup() {
  ['staff', 'exp', 'pay'].forEach(sheet_);
  var st = sheet_('set');
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 4, 2).setValues([
      ['Пароль руководителя', ''],
      ['Пароль бухгалтера', ''],
      ['Категории', DEFAULT_CATS],
      ['', '']
    ]);
  }
}
