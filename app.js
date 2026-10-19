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
let inflight = 0;
async function api(action, payload) {
  if (action !== 'login' && action !== 'roleLogin') { inflight++; try { return await apiRaw(action, payload) } finally { inflight-- } }
  return apiRaw(action, payload);
}
async function apiRaw(action, payload) {
  if (!API) throw new Error('Приложение ещё не подключено к таблице (нет адреса в config.js).');
  const r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ action }, payload || {})) });
  const j = await r.json();
  if (!j.ok) throw new Error(j.error || 'Ошибка сервера');
  return j;
}
async function load(quiet) {
  if (!API) { notice('Приложение ещё не подключено к таблице. Впишите адрес веб-приложения в config.js.'); ready = true; render(); return }
  try {
    const r = await fetch(API + '?action=data&t=' + Date.now());
    const j = await r.json();
    if (!j.ok) throw new Error(j.error);
    const d = j.data;
    allStaff = d.staff; staff = allStaff.filter(s => !s.hidden);
    cats = [...new Set((d.cats.length ? d.cats : DEFAULT_CATS).map(normCat))];
    if (inflight > 0) return;              // идёт сохранение — свежие данные возьмём в следующий раз
    expenses = d.expenses.map(e => Object.assign(e, { cat: normCat(e.cat) })); payments = d.payments;
    mergePending();
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
const isOwner = () => role === 'owner' || role === 'acc'; // бухгалтер — полные права
const canPay = () => role === 'owner' || role === 'acc';
function canDel(e) { if (isOwner()) return true; const c = current(); return !!c && e.staff === c.id && !e.rej }
const counts = e => !e.rej;   // расход «Не возмещается» не идёт в выдачу, ведомость и отчёт
const canUndoPay = p => isOwner();
function authFor(staffId) { return isOwner() ? { rolePass } : { staff: staffId, pin: pinOf(staffId) } }
const paidBy = id => payments.filter(p => p.staff === id).reduce((a, p) => a + p.amount, 0);
const spentBy = id => expenses.filter(e => e.staff === id && counts(e)).reduce((a, e) => a + e.amount, 0);
const nameOf = id => { const s = allStaff.find(x => x.id === id); return s ? s.name : '—' };

// ---- отрисовка ----
// цвет категории: [выбрана, фон, текст] — спокойные оттенки
const CAT_COLORS = {
  '주유비': ['#c26a1b', '#fbeedf', '#8a4a10'],
  '식대': ['#2f8a57', '#e4f3ea', '#1f6b41'],
  '교통비': ['#2f63c8', '#e5edfb', '#1f4a9c'],
  '하이패스': ['#2a8a99', '#e2f3f5', '#1d6470'],
  '택배비': ['#a8762a', '#f6eedf', '#7a5418'],
  '차량정비': ['#b04a3c', '#f8e7e4', '#86352a'],
  '접대': ['#8a5cc2', '#efe8f8', '#5f3b8f'],
  '사무용품': ['#5b6fa8', '#e9ecf6', '#3f4f7d'],
  '기타': ['#6b6f78', '#eceef1', '#4a4d54']
};
// перевод категорий — только для экрана (в Excel и отчётах остаётся корейский)
const CAT_RU = { '주유비': 'Бензин', '식대': 'Еда', '교통비': 'Транспорт', '하이패스': 'Платная дорога', '택배비': 'Почта, доставка',
  '차량정비': 'Ремонт авто', '접대': 'Представительские', '사무용품': 'Канцтовары', '기타': 'Прочее' };
// старые русские названия → корейские
const CAT_ALIAS = { 'Бензин': '주유비', 'Еда': '식대', 'Транспорт': '교통비', 'Прочее': '기타' };
const DEFAULT_CATS = ['주유비', '식대', '교통비', '하이패스', '택배비', '차량정비', '접대', '사무용품', '기타'];
const normCat = c => CAT_ALIAS[c] || c;
const catLabel = c => CAT_RU[c] ? `${c} · ${CAT_RU[c]}` : c;
const CAT_FALLBACK = [['#8a5cc2', '#efe8f8', '#5f3b8f'], ['#b8456b', '#f8e6ec', '#8a2f4f']];
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
  const cs = current();
  $('meLine').hidden = !cs || isOwner();
  if (cs) $('meName').textContent = `Вы вошли как ${cs.name}`;
  $('pinRow').hidden = !need;
  if (!need) $('pinErr').hidden = true;
  $('cats').style.gridTemplateColumns = `repeat(${cats.length > 4 ? 3 : (cats.length || 1)},minmax(0,1fr))`;
  $('cats').innerHTML = cats.map((c, i) => { const k = CAT_COLORS[c] || CAT_FALLBACK[i % CAT_FALLBACK.length];
    return `<button type="button" data-cat="${esc(c)}" aria-pressed="${c === cat}" style="--c:${k[0]};--c-soft:${k[1]};--c-ink:${k[2]}"><b>${esc(c)}</b>${CAT_RU[c] ? `<small>${esc(CAT_RU[c])}</small>` : ''}</button>` }).join('');
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
    const rj = e.rej ? `<span class="rejtag">Не возмещается: ${esc(e.rej)}</span>` : '';
    const rb = isOwner() && !e.pending ? (e.rej ? `<button type="button" class="linkbtn rejbtn" data-unrej="${esc(e.id)}">Вернуть к выдаче</button>` : `<button type="button" class="linkbtn rejbtn" data-rej="${esc(e.id)}">Не возмещать</button>`) : '';
    return `<div class="li${e.rej ? ' rej' : ''}" data-row="${esc(e.id)}"><span style="min-width:0"><b>${esc([e.place, e.note].filter(Boolean).join(' — ') || e.cat)}</b><small>${dmy(dateOf(e))} · ${esc(nameOf(e.staff))} · ${esc(catLabel(e.cat))}${e.pending ? ' · <i>сохраняется…</i>' : ''}</small>${rj || rb ? `<small class="rejline">${rj}${rb}</small>` : ''}</span>
      <span class="a">−${fmt(e.amount)}</span>${canDel(e) && !e.pending ? `<button class="del" type="button" data-del="${esc(e.id)}" aria-label="Удалить">×</button>` : '<span></span>'}</div>`
  }).join('') : '<div class="empty">Расходов пока нет.</div>';
}
function renderRole() {
  $('tools').hidden = !canPay();
  $('sheetLink').hidden = true;
  $('roleLine').innerHTML = role
    ? `<span>Вы вошли как бухгалтер</span><button type="button" class="linkbtn" id="roleOut">Выйти</button>`
    : `<button type="button" class="linkbtn" id="roleOpen">Вход для бухгалтера</button>`;
}
function render() { renderForm(); renderSummary(); renderHistory(); renderRole(); renderTaxYears() }

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
  catch (e2) { err.textContent = 'Неверный пароль. Если забыли — спросите у бухгалтера.'; err.hidden = false; $('pinIn').select() }
  finally { $('pinGo').disabled = false }
}
$('pinGo').addEventListener('click', tryPin);
// выход сотрудника: забыть пароль на этом телефоне
$('meOut').addEventListener('click', () => {
  if (me) ls.del('pin_' + me);
  me = ''; ls.set('me', ''); $('pinIn').value = ''; msg('', ''); render();
});
$('pinIn').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); tryPin() } });

let busy = false;
// ---- очередь отправки: расход появляется сразу, в таблицу уходит в фоне ----
function getQueue() { try { return JSON.parse(ls.get('queue') || '[]') } catch (e) { return [] } }
function setQueue(q) { ls.set('queue', JSON.stringify(q)) }
function mergePending() {
  for (const q of getQueue()) if (!expenses.some(e => e.id === q.payload.clientId)) expenses.push(Object.assign({}, q.local, { pending: true }));
}
let flushing = false;
async function flush() {
  if (flushing) return; flushing = true;
  try {
    let q = getQueue();
    while (q.length) {
      const item = q[0];
      try {
        await api('add', item.payload);
        const ex = expenses.find(e => e.id === item.payload.clientId); if (ex) delete ex.pending;
        q = getQueue().filter(x => x.payload.clientId !== item.payload.clientId); setQueue(q);
        if ($('msg').dataset.pid === item.payload.clientId) msg($('msg').textContent.replace(' · сохраняется…', '') + ' ✓', 'ok');
        render();
      } catch (err) {
        if (err instanceof TypeError || /Failed to fetch|NetworkError|Load failed/i.test(err.message)) {
          msg('Нет связи — расход сохранён в телефоне и отправится автоматически, когда появится интернет.', 'err');
          break;   // нет интернета: попробуем позже
        }
        // сервер отказал (пароль и т. п.) — убираем из очереди и с экрана
        q = getQueue().filter(x => x.payload.clientId !== item.payload.clientId); setQueue(q);
        expenses = expenses.filter(e => e.id !== item.payload.clientId);
        if (/пароль/i.test(err.message)) ls.del('pin_' + item.payload.staff);
        msg(`Не сохранилось (${fmt(item.local.amount)}): ${err.message}`, 'err'); render();
      }
    }
  } finally { flushing = false }
}
window.addEventListener('online', flush);

$('form').addEventListener('submit', async e => {
  e.preventDefault(); if (busy) return;
  const who = current(), a = parse(amt.value), date = $('date').value || isoDay(0);
  if (!me) { msg('Выберите сотрудника.', 'err'); $('who').focus(); return }
  if (!who) { msg('Введите пароль сотрудника.', 'err'); $('pinIn').focus(); return }
  if (!a) { msg('Введите сумму.', 'err'); amt.focus(); return }
  if (a > 100000000) { msg('Сумма слишком большая — проверьте, нет ли лишних нулей.', 'err'); return }
  if (date > isoDay(0)) { msg('Дата не может быть в будущем.', 'err'); return }
  if (!$('place').value.trim()) { msg('Укажите, где потратили.', 'err'); $('place').focus(); return }
  if (!$('note').value.trim()) { msg('Напишите комментарий.', 'err'); $('note').focus(); return }
  const note = $('note').value.trim(), place = $('place').value.trim();
  const clientId = 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const local = { id: clientId, date, staff: who.id, cat, amount: a, note, place, created: Date.now() };
  const payload = Object.assign({ cat, amount: a, date, note, place, clientId }, authFor(who.id), { staff: who.id });
  setQueue(getQueue().concat([{ payload, local }]));
  expenses.push(Object.assign({}, local, { pending: true }));
  msg(`Списано ${fmt(a)} · ${cat} · ${who.name} · ${dmy(date)} · сохраняется…`, 'ok'); $('msg').dataset.pid = clientId;
  amt.value = ''; $('note').value = ''; $('place').value = ''; render();
  flush();
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
    // сразу показываем результат, сервер подтверждает в фоне
    const tmp = { id: 'tmp' + Date.now(), date: isoDay(0), staff: s.id, amount: amtv, by: 'Бухгалтер', created: Date.now(), pending: true };
    payments.push(tmp); render();
    pm.className = 'msg ok'; pm.textContent = `Выдано: ${s.name} — ${fmt(amtv)} · сохраняется…`;
    try {
      const r = await api('pay', { rolePass, staff: s.id, amount: amtv });
      tmp.id = r.id; delete tmp.pending; render();
      pm.className = 'msg ok'; pm.innerHTML = ''; pm.append(`Выдано: ${s.name} — ${fmt(amtv)} ✓ `);
      const ub = document.createElement('button'); ub.type = 'button'; ub.className = 'undo'; ub.textContent = 'Отменить'; ub.dataset.delpay = r.id; ub.dataset.now = '1'; pm.append(ub);
    } catch (err) { payments = payments.filter(x => x !== tmp); pm.className = 'msg err'; pm.textContent = 'Не сохранилось: ' + err.message; render() }
    return;
  }

  const dp = e.target.closest('[data-delpay]');
  if (dp) {
    const p = payments.find(x => x.id === dp.dataset.delpay); if (!p || !canUndoPay(p)) return;
    if (!dp.dataset.now && !dp.classList.contains('armed')) {
      dp.classList.add('armed'); dp.textContent = 'Точно отменить?';
      setTimeout(() => { if (dp.isConnected) { dp.classList.remove('armed'); dp.textContent = 'Отменить' } }, 4000); return
    }
    if (p.pending) return;
    payments = payments.filter(x => x.id !== p.id); $('payMsg').className = 'msg'; $('payMsg').textContent = 'Выдача отменена.'; render();
    try { await api('unpay', { rolePass, id: p.id }) }
    catch (err) { payments.push(p); $('payMsg').className = 'msg err'; $('payMsg').textContent = 'Не удалось отменить: ' + err.message; render() }
    return;
  }

  // «Не возмещать»: бухгалтер указывает причину, расход остаётся в истории, но не идёт в выдачу
  const rq = e.target.closest('[data-rej]');
  if (rq) {
    if (!isOwner()) return;
    document.querySelectorAll('.rejform').forEach(f => f.remove());
    const row = rq.closest('.li'), f = document.createElement('form');
    f.className = 'rejform'; f.dataset.id = rq.dataset.rej; f.noValidate = true;
    f.innerHTML = `<input type="text" maxlength="80" value="нет чека" aria-label="Причина"><button type="submit" class="tubtn">Не возмещать</button><button type="button" class="linkbtn" data-rejcancel="1">Отмена</button>`;
    row.after(f); f.querySelector('input').select();
    return;
  }
  if (e.target.closest('[data-rejcancel]')) { e.target.closest('.rejform').remove(); return }
  const ur = e.target.closest('[data-unrej]');
  if (ur) { setRej(ur.dataset.unrej, ''); return }

  const d = e.target.closest('[data-del]');
  if (d) {
    const ex = expenses.find(x => x.id === d.dataset.del); if (!ex || !canDel(ex)) return;
    if (!d.classList.contains('armed')) { d.classList.add('armed'); d.textContent = 'Удалить?'; setTimeout(() => { if (d.isConnected) { d.classList.remove('armed'); d.textContent = '×' } }, 3500); return }
    expenses = expenses.filter(x => x.id !== ex.id); render();
    try { await api('del', Object.assign({ id: ex.id }, authFor(ex.staff))) }
    catch (err) { expenses.push(ex); msg('Не удалось удалить: ' + err.message, 'err'); render() }
  }
});

document.addEventListener('submit', e => {
  const f = e.target.closest('.rejform'); if (!f) return;
  e.preventDefault();
  setRej(f.dataset.id, f.querySelector('input').value.trim() || 'не возмещается');
});
async function setRej(id, reason) {
  const ex = expenses.find(x => x.id === id); if (!ex || !isOwner()) return;
  const old = ex.rej || '';
  ex.rej = reason; render();
  try { const r = await api(reason ? 'reject' : 'unreject', { rolePass, id, reason }); ex.rej = r.rej || ''; render() }
  catch (err) { ex.rej = old; msg('Не сохранилось: ' + err.message, 'err'); render() }
}

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
    expenses: expenses.slice().sort((a, b) => dateOf(a).localeCompare(dateOf(b)) || (a.created || 0) - (b.created || 0)).map(e => ({ date: ser(dateOf(e)), name: nameOf(e.staff), cat: e.cat, amount: e.amount, note: e.note || '', rej: e.rej || '' })),
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

// ---- Отчёт для налоговой: 현금 지출 현황 по году ----
function renderTaxYears() {
  const ys = [...new Set(expenses.map(e => dateOf(e).slice(0, 4)).concat([isoDay(0).slice(0, 4)]))].sort().reverse();
  const sel = $('taxYear'), cur = sel.value || ys[0];
  if (sel.dataset.sig !== ys.join()) { sel.innerHTML = ys.map(y => `<option value="${y}">${y}</option>`).join(''); sel.dataset.sig = ys.join('') }
  sel.value = ys.includes(cur) ? cur : ys[0];
}
async function exportTax() {
  const lm = (t, k) => { $('lgMsg').className = 'msg ' + k; $('lgMsg').textContent = t };
  if (!canPay()) return;
  if (typeof XLSX === 'undefined') { lm('Не загрузился модуль Excel. Обновите страницу.', 'err'); return }
  await load(true);
  const year = $('taxYear').value || isoDay(0).slice(0, 4);
  const ser = iso => { const [y, m, d] = iso.split('-').map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5 };
  const rows = expenses.filter(e => counts(e) && dateOf(e).startsWith(year + '-'))
    .sort((a, b) => dateOf(a).localeCompare(dateOf(b)) || (a.created || 0) - (b.created || 0))
    .map(e => ({ date: ser(dateOf(e)), place: e.place || '', detail: [e.cat, e.note].filter(Boolean).join(' - '), amount: e.amount, memo: nameOf(e.staff) }));
  try {
    const buf = XLSX.write(buildTaxBook(XLSX, year, rows), { bookType: 'xlsx', type: 'array' });
    await saveFile(buf, `현금지출현황_${year}.xlsx`);
    lm(`Отчёт за ${year} готов: ${rows.length} записей.`, 'ok');
  } catch (err) { lm('Не удалось собрать отчёт.', 'err') }
}
async function saveFile(buf, name) {
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const file = new File([blob], name, { type: blob.type });
  if (navigator.canShare && navigator.canShare({ files: [file] }) && /iPhone|iPad|Android/i.test(navigator.userAgent)) {
    await navigator.share({ files: [file], title: name }).catch(() => {});
  } else {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove() }, 2000);
  }
}
$('dlTax').addEventListener('click', exportTax);

// ---- запуск и автообновление ----
render(); load().then(flush);
setInterval(() => {
  if (document.hidden) return;
  $('date').max = isoDay(0);
  const ae = document.activeElement;
  if (ae && /^(INPUT|SELECT)$/.test(ae.tagName)) return;
  if (document.querySelector('.armed')) return;
  load(true);
}, 30000);
document.addEventListener('visibilitychange', () => { if (!document.hidden && Date.now() - lastLoad > 10000) load(true) });
