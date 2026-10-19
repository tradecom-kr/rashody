/* Расходы команды — клиент. Данные и проверки паролей — на сервере Google Apps Script (config.js → API_URL). */
'use strict';
const API = (window.APP_CONFIG && window.APP_CONFIG.API_URL) || '';
const SHEET_URL = (window.APP_CONFIG && window.APP_CONFIG.SHEET_URL) || '';

const $ = id => document.getElementById(id);
const fmt = n => Math.round(n).toLocaleString('ru-RU') + ' ₩';
const nf = n => Math.round(n).toLocaleString('en-US');
const parse = v => Number(String(v).replace(/[^\d]/g, '')) || 0;
const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const ls = {
  get(k) { try { return localStorage.getItem(k) } catch (e) { return null } },
  set(k, v) { try { localStorage.setItem(k, v) } catch (e) {} },
  del(k) { try { localStorage.removeItem(k) } catch (e) {} }
};

// ---- состояние ----
let allStaff = [], staff = [], cats = [], expenses = [], payments = [];
let me = ls.get('me') || '', cat = ls.get('cat') || '', hWho = 'all';
let role = ls.get('role') || '', rolePass = ls.get('rolePass') || '';
let ready = false, lastLoad = 0;

// ---- даты ----
function isoDay(offset) { const d = new Date(); d.setDate(d.getDate() + (offset || 0)); return isoOf(d.getTime()) }
const isoOf = ms => { const d = new Date(ms); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}` };
const dm = ms => { const d = new Date(ms); return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}` };
const dmy = iso => iso ? iso.split('-').join('.') : '';
const ymd = ms => isoOf(ms).split('-').join('.');
const dateOf = e => e.date || isoOf(e.created || Date.now());

// ---- сервер ----
async function api(action, payload) {
  if (!API) throw new Error('Приложение ещё не подключено к таблице (нет адреса в config.js).');
  const r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ action }, payload || {})) });
  const j = await r.json();
  if (!j.ok) throw new Error(j.error || 'Ошибка сервера');
  return j;
}
async function load(quiet) {
  if (!API) { notice('Приложение ещё не подключено к таблице. Руководителю: впишите адрес веб-приложения в config.js.'); ready = true; render(); return }
  try {
    const r = await fetch(API + '?action=data&t=' + Date.now());
    const j = await r.json();
    if (!j.ok) throw new Error(j.error);
    const d = j.data;
    allStaff = d.staff; staff = allStaff.filter(s => !s.hidden);
    cats = d.cats.length ? d.cats : ['Бензин', 'Еда', 'Транспорт', 'Прочее'];
    expenses = d.expenses; payments = d.payments;
    if (me && !staff.some(s => s.id === me)) { me = ''; ls.set('me', '') }
    if (!cats.includes(cat)) cat = cats[0];
    ready = true; lastLoad = Date.now(); $('notice').hidden = true;
    render();
  } catch (err) {
    if (!quiet) notice('Нет связи с таблицей. Проверьте интернет и обновите страницу.');
    ready = true; render();
  }
}
function notice(t) { $('notice').textContent = t; $('notice').hidden = false }

// ---- кто вошёл ----
const pinOf = id => ls.get('pin_' + id) || '';
function isVerified(id) {
  if (role === 'owner') return true;
  const s = staff.find(x => x.id === id); if (!s) return false;
  return !s.hasPin || !!pinOf(id);
}
function current() { const s = staff.find(x => x.id === me); return s && isVerified(s.id) ? s : null }
const isOwner = () => role === 'owner';
const canPay = () => role === 'owner' || role === 'acc';
function canDel(e) { if (isOwner()) return true; const c = current(); return !!c && e.staff === c.id }
const canUndoPay = p => isOwner() || (role === 'acc' && p.by === 'Бухгалтер');
function authFor(staffId) { return isOwner() ? { rolePass } : { staff: staffId, pin: pinOf(staffId) } }
const paidBy = id => payments.filter(p => p.staff === id).reduce((a, p) => a + p.amount, 0);
const spentBy = id => expenses.filter(e => e.staff === id).reduce((a, e) => a + e.amount, 0);
const nameOf = id => { const s = allStaff.find(x => x.id === id); return s ? s.name : '—' };

// ---- отрисовка ----
const AVC = ['#031b95', '#404040', '#2f4fd0', '#6b6e75', '#14309f', '#55575c', '#4561d8'];
const initials = n => String(n).split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();

function renderForm() {
  const sel = $('who');
  const sig = staff.map(s => s.id + ':' + s.name).join('|');
  if (sel.dataset.sig !== sig) {           // пересобираем список, только когда сотрудники изменились
    sel.innerHTML = `<option value="">Кто вы?</option>` + staff.map(s => `<option value="${esc(s.id)}">${esc(s.name)}</option>`).join('');
    sel.dataset.sig = sig;
  }
  if (sel.value !== me) sel.value = me;
  const need = me && !isVerified(me);
  $('pinRow').hidden = !need;
  if (!need) $('pinErr').hidden = true;
  $('cats').style.gridTemplateColumns = `repeat(${Math.min(cats.length, 4) || 1},minmax(0,1fr))`;
  $('cats').innerHTML = cats.map(c => `<button type="button" data-cat="${esc(c)}" aria-pressed="${c === cat}">${esc(c)}</button>`).join('');
}
function renderSummary() {
  $('wkSub').textContent = canPay() ? 'После перевода нажмите «Выдать» и подтвердите — сумма к выдаче станет 0.' : '«К выдаче» — сколько ещё не возмещено сотруднику.';
  $('updated').textContent = lastLoad ? 'обновлено ' + new Date(lastLoad).toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' }) : 'загрузка…';
  $('sum').innerHTML = staff.map((s, i) => {
    const sp = spentBy(s.id), pd = paidBy(s.id), due = sp - pd, rest = s.budget - due;
    const pct = s.budget ? Math.min(100, Math.max(0, due / s.budget * 100)) : 0;
    return `<article class="pc">
      <div class="pc-top"><span class="av" style="--c:${AVC[i % AVC.length]}">${esc(initials(s.name))}</span><span class="pc-name">${esc(s.name)}</span></div>
      <div><div class="due-l">К выдаче</div><div class="due ${due ? '' : 'zero'}">${(due < 0 ? '−' : '') + nf(Math.abs(due))}</div></div>
      ${s.budget ? `<div class="meter ${rest < 0 ? 'over' : ''}"><i style="width:${rest < 0 ? 100 : pct}%"></i></div>` : ''}
      <div class="rest ${rest < 0 ? 'neg' : ''}">${s.budget ? `Остаток <b>${(rest < 0 ? '−' : '') + nf(Math.abs(rest))}</b> из ${nf(s.budget)}` : 'Без бюджета'}</div>
      ${canPay() ? `<button type="button" class="paybtn" data-pay="${esc(s.id)}" data-amt="${due}"${due > 0 ? '' : ' disabled'}>Выдать</button>` : ''}
    </article>`
  }).join('');
  $('payMsg').hidden = !canPay();
}
function renderHistory() {
  const sel = $('hWho');
  const sig = staff.map(s => s.id + ':' + s.name).join('|');
  if (sel.dataset.sig !== sig) { sel.innerHTML = `<option value="all">Все сотрудники</option>` + staff.map(s => `<option value="${esc(s.id)}">${esc(s.name)}</option>`).join(''); sel.dataset.sig = sig }
  if (sel.value !== hWho) sel.value = hWho;
  if (!ready) { $('log').innerHTML = '<div class="empty">Загрузка…</div>'; return }
  const list = [...expenses, ...payments.map(p => Object.assign({}, p, { pmt: true }))]
    .filter(e => hWho === 'all' || e.staff === hWho)
    .sort((a, b) => dateOf(b).localeCompare(dateOf(a)) || (b.created || 0) - (a.created || 0)).slice(0, 300);
  $('log').innerHTML = list.length ? list.map(e => {
    if (e.pmt) return `<div class="li pmt"><span style="min-width:0"><b>Выдано</b><small>${dmy(dateOf(e))} · ${esc(nameOf(e.staff))} · ${esc(e.by || 'выдача')}</small></span>
      <span class="a">+${fmt(e.amount)}</span>${canUndoPay(e) ? `<button class="undo" type="button" data-delpay="${esc(e.id)}">Отменить</button>` : '<span></span>'}</div>`;
    return `<div class="li"><span style="min-width:0"><b>${esc(e.note || e.cat)}</b><small>${dmy(dateOf(e))} · ${esc(nameOf(e.staff))} · ${esc(e.cat)}</small></span>
      <span class="a">−${fmt(e.amount)}</span>${canDel(e) ? `<button class="del" type="button" data-del="${esc(e.id)}" aria-label="Удалить">×</button>` : '<span></span>'}</div>`
  }).join('') : '<div class="empty">Расходов пока нет.</div>';
}
function renderRole() {
  $('tools').hidden = !canPay();
  if (SHEET_URL && isOwner()) { $('sheetLink').href = SHEET_URL; $('sheetLink').hidden = false } else $('sheetLink').hidden = true;
  $('roleLine').innerHTML = role
    ? `<span>Вы вошли как ${role === 'owner' ? 'руководитель' : 'бухгалтер'}</span><button type="button" class="linkbtn" id="roleOut">Выйти</button>`
    : `<button type="button" class="linkbtn" id="roleOpen">Вход для руководителя или бухгалтера</button>`;
}
function render() { renderForm(); renderSummary(); renderHistory(); renderRole() }

// ---- ввод расхода ----
const amt = $('amount');
amt.addEventListener('input', () => { const n = parse(amt.value); amt.value = n ? n.toLocaleString('ru-RU') : '' });
$('date').value = isoDay(0); $('date').max = isoDay(0);
function msg(t, k) { $('msg').textContent = t; $('msg').className = 'msg ' + (k || '') }

$('who').addEventListener('change', e => {
  me = e.target.value; ls.set('me', me); $('pinIn').value = ''; $('pinErr').hidden = true; render();
  if (me && !isVerified(me)) $('pinIn').focus();
});
async function tryPin() {
  const s = staff.find(x => x.id === me); if (!s) return;
  const v = $('pinIn').value.trim(), err = $('pinErr');
  if (!v) { err.textContent = 'Введите пароль.'; err.hidden = false; return }
  $('pinGo').disabled = true;
  try { await api('login', { staff: s.id, pin: v }); ls.set('pin_' + s.id, v); $('pinIn').value = ''; render() }
  catch (e2) { err.textContent = 'Неверный пароль. Если забыли — спросите у руководителя.'; err.hidden = false; $('pinIn').select() }
  finally { $('pinGo').disabled = false }
}
$('pinGo').addEventListener('click', tryPin);
$('pinIn').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); tryPin() } });

let busy = false;
$('form').addEventListener('submit', async e => {
  e.preventDefault(); if (busy) return;
  const who = current(), a = parse(amt.value), date = $('date').value || isoDay(0);
  if (!me) { msg('Выберите сотрудника.', 'err'); $('who').focus(); return }
  if (!who) { msg('Введите пароль сотрудника.', 'err'); $('pinIn').focus(); return }
  if (!a) { msg('Введите сумму.', 'err'); amt.focus(); return }
  if (a > 100000000) { msg('Сумма слишком большая — проверьте, нет ли лишних нулей.', 'err'); return }
  if (date > isoDay(0)) { msg('Дата не может быть в будущем.', 'err'); return }
  busy = true; $('go').disabled = true; $('go').textContent = 'Сохраняю…';
  try {
    const note = $('note').value.trim();
    const r = await api('add', Object.assign({ cat, amount: a, date, note }, authFor(who.id), { staff: who.id }));
    expenses.push({ id: r.id, date, staff: who.id, cat, amount: a, note, created: Date.now() });
    msg(`Списано ${fmt(a)} · ${cat} · ${who.name} · ${dmy(date)}`, 'ok');
    amt.value = ''; $('note').value = ''; render(); load(true);
  } catch (err) {
    if (/пароль/i.test(err.message)) { ls.del('pin_' + who.id); render() }
    msg(err.message || 'Не сохранилось. Попробуйте ещё раз.', 'err');
  } finally { busy = false; $('go').disabled = false; $('go').textContent = 'Списать' }
});

// ---- клики: категории, выдачи, удаление ----
function disarmPay() { document.querySelectorAll('.paybtn.armed').forEach(b => { b.classList.remove('armed'); b.textContent = 'Выдать' }); document.querySelectorAll('.paycancel').forEach(b => b.remove()) }
document.addEventListener('click', async e => {
  const c = e.target.closest('[data-cat]'); if (c) { cat = c.dataset.cat; ls.set('cat', cat); renderForm(); return }
  if (e.target.closest('#roleOpen')) { $('roleBox').hidden = false; $('rolePass').focus(); return }
  if (e.target.closest('#roleOut')) { role = ''; rolePass = ''; ls.del('role'); ls.del('rolePass'); render(); return }
  if (e.target.closest('[data-paycancel]')) { disarmPay(); $('payMsg').textContent = ''; return }

  const pb = e.target.closest('[data-pay]');
  if (pb) {
    if (!canPay()) return;
    const s = staff.find(x => x.id === pb.dataset.pay), amtv = Number(pb.dataset.amt); if (!s || !(amtv > 0)) return;
    const pm = $('payMsg');
    if (!pb.classList.contains('armed')) {
      disarmPay();
      pb.classList.add('armed'); pb.textContent = 'Да, выдать ' + nf(amtv);
      const cb = document.createElement('button'); cb.type = 'button'; cb.className = 'paycancel'; cb.textContent = 'Отмена'; cb.dataset.paycancel = '1'; pb.after(cb);
      pm.className = 'msg'; pm.textContent = `Выдать ${s.name} ${fmt(amtv)}?`;
      return;
    }
    pb.disabled = true; pb.textContent = 'Сохраняю…';
    try {
      const r = await api('pay', { rolePass, staff: s.id, amount: amtv });
      payments.push({ id: r.id, date: isoDay(0), staff: s.id, amount: amtv, by: isOwner() ? 'Руководитель' : 'Бухгалтер', created: Date.now() });
      render();
      pm.className = 'msg ok'; pm.innerHTML = ''; pm.append(`Выдано: ${s.name} — ${fmt(amtv)}. `);
      const ub = document.createElement('button'); ub.type = 'button'; ub.className = 'undo'; ub.textContent = 'Отменить'; ub.dataset.delpay = r.id; ub.dataset.now = '1'; pm.append(ub);
      load(true);
    } catch (err) { pm.className = 'msg err'; pm.textContent = err.message; render() }
    return;
  }

  const dp = e.target.closest('[data-delpay]');
  if (dp) {
    const p = payments.find(x => x.id === dp.dataset.delpay); if (!p || !canUndoPay(p)) return;
    if (!dp.dataset.now && !dp.classList.contains('armed')) {
      dp.classList.add('armed'); dp.textContent = 'Точно отменить?';
      setTimeout(() => { if (dp.isConnected) { dp.classList.remove('armed'); dp.textContent = 'Отменить' } }, 4000); return
    }
    dp.disabled = true;
    try { await api('unpay', { rolePass, id: p.id }); payments = payments.filter(x => x.id !== p.id); $('payMsg').className = 'msg'; $('payMsg').textContent = 'Выдача отменена.'; render(); load(true) }
    catch (err) { $('payMsg').className = 'msg err'; $('payMsg').textContent = err.message; dp.disabled = false }
    return;
  }

  const d = e.target.closest('[data-del]');
  if (d) {
    const ex = expenses.find(x => x.id === d.dataset.del); if (!ex || !canDel(ex)) return;
    if (!d.classList.contains('armed')) { d.classList.add('armed'); d.textContent = 'Удалить?'; setTimeout(() => { if (d.isConnected) { d.classList.remove('armed'); d.textContent = '×' } }, 3500); return }
    d.disabled = true;
    try { await api('del', Object.assign({ id: ex.id }, authFor(ex.staff))); expenses = expenses.filter(x => x.id !== ex.id); render(); load(true) }
    catch (err) { msg(err.message, 'err'); d.disabled = false }
  }
});

$('roleBox').addEventListener('submit', async e => {
  e.preventDefault();
  const v = $('rolePass').value.trim(), err = $('roleErr');
  if (!v) return;
  try {
    const r = await api('roleLogin', { rolePass: v });
    role = r.role; rolePass = v; ls.set('role', role); ls.set('rolePass', v);
    $('rolePass').value = ''; $('roleBox').hidden = true; err.hidden = true; render();
  } catch (e2) { err.textContent = 'Неверный пароль.'; err.hidden = false }
});
$('hWho').addEventListener('change', e => { hWho = e.target.value; renderHistory() });

// ---- Excel ----
async function exportXlsx() {
  const lm = (t, k) => { $('lgMsg').className = 'msg ' + k; $('lgMsg').textContent = t };
  if (!canPay()) return;
  if (typeof XLSX === 'undefined') { lm('Не загрузился модуль Excel. Обновите страницу.', 'err'); return }
  await load(true);
  const ser = iso => { if (!iso) return 0; const [y, m, d] = iso.split('-').map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5 };
  const data = {
    today: dmy(isoDay(0)),
    staff: staff.map(s => { const sp = spentBy(s.id), pd = paidBy(s.id); return { name: s.name, budget: s.budget, spent: sp, paid: pd, due: sp - pd } }),
    expenses: expenses.slice().sort((a, b) => dateOf(a).localeCompare(dateOf(b)) || (a.created || 0) - (b.created || 0)).map(e => ({ date: ser(dateOf(e)), name: nameOf(e.staff), cat: e.cat, amount: e.amount, note: e.note || '' })),
    payments: payments.slice().sort((a, b) => (a.created || 0) - (b.created || 0)).map(p => ({ date: ser(p.date), name: nameOf(p.staff), amount: p.amount }))
  };
  data.weeks = weeklyLedger();
  try {
    const buf = XLSX.write(buildBook(XLSX, data), { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const name = `Расходы команды ${dmy(isoDay(0))}.xlsx`;
    const file = new File([blob], name, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] }) && /iPhone|iPad|Android/i.test(navigator.userAgent)) {
      await navigator.share({ files: [file], title: name }).catch(() => {});
    } else {
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove() }, 2000);
    }
    lm('Файл готов.', 'ok');
  } catch (err) { lm('Не удалось собрать файл.', 'err') }
}
$('dlLedger').addEventListener('click', exportXlsx);

// ---- запуск и автообновление ----
render(); load();
setInterval(() => {
  if (document.hidden) return;
  $('date').max = isoDay(0);
  const ae = document.activeElement;
  if (ae && /^(INPUT|SELECT)$/.test(ae.tagName)) return;
  if (document.querySelector('.armed')) return;
  load(true);
}, 30000);
document.addEventListener('visibilitychange', () => { if (!document.hidden && Date.now() - lastLoad > 10000) load(true) });
