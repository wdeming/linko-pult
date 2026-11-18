/* Учёт долгов: реестр, карточка долга, выплаты и начисления, журнал. Данные — /api/debts (база пульта, Odoo не используется). */
(() => {
  const root = document.getElementById('debts'); if (!root) return;
  const $ = (id) => document.getElementById(id);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const nf0 = new Intl.NumberFormat('ru-RU', {maximumFractionDigits: 0}), nf2 = new Intl.NumberFormat('ru-RU', {maximumFractionDigits: 2});
  const money = (a, cur) => cur === 'USD' ? '$' + nf2.format(a) : nf0.format(a) + ' сум';
  const ru = (d) => d ? d.slice(8, 10) + '.' + d.slice(5, 7) + '.' + d.slice(0, 4) : '—';
  const when = (t) => t ? new Date(t).toLocaleString('ru-RU', {dateStyle: 'short', timeStyle: 'short'}) : '';
  const ST = {open: ['Открыт', ''], overdue: ['Просрочен', 'bad'], closed: ['Погашен', 'ok'], void: ['Аннулирован', 'void']};
  let S = null, filter = 'active', q = '', cardId = null;
  const can = () => S && S.can === 'edit';
  async function api(url, body) {
    const r = await fetch(url, body ? {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)} : {});
    if (r.status === 401) { location.href = '/login'; throw new Error('Нужен вход'); }
    const j = await r.json().catch(() => ({})); if (!r.ok) throw new Error(j.error || 'Ошибка сервера'); return j;
  }
  async function load() {
    try { S = await api('/api/debts'); }
    catch (e) {
      if (/доступа/.test(e.message)) { root.remove(); document.querySelector('.pl-tab[data-tab="debts"]')?.remove(); return; }
      $('db-table').replaceChildren(el('p', 'pl-file', 'Не удалось загрузить долги: ' + e.message)); return;
    }
    $('db-new').hidden = !can(); render(); if (cardId != null && $('db-card').open) card(cardId);
  }
  function tileBox(label, value, sub, tone) {
    const t = el('div', 'pl-tile db-tile' + (tone ? ' ' + tone : '')); t.append(el('div', 'pl-tl', label), el('div', 'pl-tv', value)); if (sub) t.append(el('div', 'db-tsub', sub)); return t;
  }
  function render() {
    const live = S.debts.filter((d) => !d.void), m = S.today.slice(0, 7);
    const left = live.reduce((a, d) => a + d.left_uzs, 0), usd = live.filter((d) => d.currency === 'USD').reduce((a, d) => a + d.left, 0);
    const paidM = live.reduce((a, d) => a + d.ops.filter((o) => !o.void && o.op === 'pay' && o.date.slice(0, 7) === m).reduce((x, o) => x + o.amount * (d.currency === 'USD' ? S.rate : 1), 0), 0);
    const takenM = live.reduce((a, d) => a + d.ops.filter((o) => !o.void && o.op === 'take' && o.date.slice(0, 7) === m).reduce((x, o) => x + o.amount * (d.currency === 'USD' ? S.rate : 1), 0), 0);
    const od = live.filter((d) => d.status === 'overdue'), open = live.filter((d) => d.status !== 'closed');
    $('db-tiles').replaceChildren(
      tileBox('Всего должны сейчас', money(left, 'UZS'), usd ? 'в т.ч. $' + nf2.format(usd) + ' по 12 000' : 'открытых долгов: ' + open.length, left > 0 ? 'bad' : 'good'),
      tileBox('Было взято всего', money(live.reduce((a, d) => a + d.taken_uzs, 0), 'UZS'), 'с учётом добавленного к долгам'),
      tileBox('Вернули всего', money(live.reduce((a, d) => a + d.paid_uzs, 0), 'UZS'), 'погашено долгов: ' + live.filter((d) => d.status === 'closed').length),
      tileBox('Вернули в этом месяце', money(paidM, 'UZS'), takenM ? 'и взяли/добавили ' + money(takenM, 'UZS') : 'новых долгов в этом месяце нет'),
      tileBox('Просрочено', money(od.reduce((a, d) => a + d.left_uzs, 0), 'UZS'), od.length ? 'долгов: ' + od.length : 'просрочек нет', od.length ? 'bad' : ''));
    let rows = S.debts.filter((d) => filter === 'all' ? true : filter === 'active' ? (d.status === 'open' || d.status === 'overdue') : filter === 'closed' ? d.status === 'closed' : d.status === 'void');
    if (q) rows = rows.filter((d) => (d.title + ' ' + d.creditor + ' ' + (d.kind || '') + ' ' + (d.note || '')).toLowerCase().includes(q));
    const order = {overdue: 0, open: 1, closed: 2, void: 3}; rows.sort((a, b) => order[a.status] - order[b.status] || (a.due || '9') .localeCompare(b.due || '9') || b.opened.localeCompare(a.opened));
    const w = el('div', 'pl-tw'), t = el('table', 'pl-tbl db-tbl'), hr = el('tr');
    ['Долг', 'Тип', 'Дата', 'Было', 'Вернули', 'Осталось', 'Срок', 'Статус'].forEach((h) => hr.append(el('th', '', h)));
    const th = el('thead'); th.append(hr); const tb = el('tbody');
    if (!rows.length) { const tr = el('tr'), td = el('td', 'db-empty', S.debts.length ? 'Ничего не найдено.' : (can() ? 'Долгов пока нет. Нажмите «Новый долг», чтобы завести первый.' : 'Долгов пока нет.')); td.colSpan = 8; tr.append(td); tb.append(tr); }
    rows.forEach((d) => {
      const tr = el('tr', 'db-row' + (d.void ? ' db-void' : '')); tr.tabIndex = 0; tr.setAttribute('role', 'button'); tr.setAttribute('aria-label', 'Открыть долг ' + d.title);
      const c1 = el('td'); c1.append(el('div', 'db-t', d.title), el('div', 'db-c', d.creditor + (d.ctype ? ' · ' + d.ctype : '')));
      const st = el('span', 'db-st ' + ST[d.status][1], ST[d.status][0]), c8 = el('td'); c8.append(st);
      tr.append(c1, el('td', 'db-k', d.kind || ''), el('td', '', ru(d.opened)), el('td', '', money(d.taken, d.currency)), el('td', '', money(d.paid, d.currency)),
        el('td', 'db-left', money(d.left, d.currency)), el('td', d.status === 'overdue' ? 'db-bad' : '', d.due ? ru(d.due) : '—'), c8);
      tr.addEventListener('click', () => card(d.id)); tr.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); card(d.id); } });
      tb.append(tr);
    });
    const vis = rows.filter((d) => !d.void);
    if (vis.length > 1) {
      const tr = el('tr', 'pl-totrow'), sum = (k) => money(vis.reduce((a, d) => a + d[k], 0), 'UZS');
      tr.append(el('td', '', 'Итого, в сумах' + (vis.some((d) => d.currency === 'USD') ? ' ($ по 12 000)' : '')), el('td'), el('td'), el('td', '', sum('taken_uzs')), el('td', '', sum('paid_uzs')), el('td', 'db-left', sum('left_uzs')), el('td'), el('td'));
      tb.append(tr);
    }
    t.append(th, tb); w.append(t); $('db-table').replaceChildren(w);
    const L = $('db-log'); L.replaceChildren();
    if (!S.log.length) L.append(el('p', 'pl-file', 'Изменений пока нет.'));
    S.log.slice(0, 100).forEach((g) => { const r = el('div', 'db-lr'); r.append(el('span', 'db-lw', when(g.at) + ' · ' + (g.who || '')), el('span', '', g.detail)); L.append(r); });
  }
  // ---------- карточка долга ----------
  function card(id) {
    const d = S.debts.find((x) => x.id === id); if (!d) return; cardId = id;
    const box = $('db-card-body'); box.replaceChildren();
    const h = el('div', 'db-ch'); const ht = el('div'); ht.append(el('h3', 'pl-modal-title', d.title), el('p', 'db-c', d.creditor + (d.ctype ? ' · ' + d.ctype : '') + (d.kind ? ' · ' + d.kind : '')));
    h.append(ht, el('span', 'db-st ' + ST[d.status][1], ST[d.status][0])); box.append(h);
    const nums = el('div', 'db-nums'), k = d.currency === 'USD';
    [['Было (взято)', d.taken], ['Вернули', d.paid], ['Осталось', d.left]].forEach(([l, v], i) => { const n = el('div', 'db-num' + (i === 2 ? ' db-numl' : '')); n.append(el('span', '', l), el('b', '', money(v, d.currency))); if (k) n.append(el('small', '', '≈ ' + money(v * S.rate, 'UZS'))); nums.append(n); });
    box.append(nums);
    const f = el('dl', 'db-facts'); const fact = (a, b) => { if (b) f.append(el('dt', '', a), el('dd', '', b)); };
    fact('Дата долга', ru(d.opened)); fact('Срок возврата', d.due ? ru(d.due) + (d.status === 'overdue' ? ' — просрочен' : '') : 'не указан'); fact('Валюта', k ? 'доллары (в итогах по 12 000 сум/$)' : 'сумы');
    fact('Комментарий', d.note); fact('Заведён', when(d.created_at) + (d.created_by ? ' · ' + d.created_by : ''));
    if (d.void) fact('Аннулирован', when(d.void_at) + ' · ' + (d.void_by || '') + '. Причина: ' + d.void_reason);
    box.append(f);
    if (can() && !d.void) {
      const a = el('div', 'db-acts');
      const b = (t, cls, fn, dis) => { const x = el('button', 'pl-btn' + (cls ? ' ' + cls : ''), t); x.type = 'button'; x.disabled = !!dis; x.addEventListener('click', fn); a.append(x); };
      b('Выплата', 'pl-main', () => opForm(d, 'pay'), d.left <= 0); b('Добавить к долгу', '', () => opForm(d, 'take')); b('Изменить', '', () => debtForm(d)); b('Аннулировать долг', 'db-danger', () => voidForm('debt', d));
      box.append(a);
    }
    box.append(el('h4', 'db-h4', 'Операции'));
    const w = el('div', 'pl-tw'), t = el('table', 'pl-tbl db-ops'), hr = el('tr');
    ['Дата', 'Операция', 'Сумма', 'Комментарий', 'Внёс', ''].forEach((x) => hr.append(el('th', '', x)));
    const th = el('thead'); th.append(hr); const tb = el('tbody');
    d.ops.forEach((o) => {
      const tr = el('tr', o.void ? 'db-void' : ''), act = el('td');
      if (o.void) act.append(el('span', 'db-vr', 'аннулирована: ' + (o.void_reason || '')));
      else if (can() && !d.void) { const x = el('button', 'db-link', 'Аннулировать'); x.type = 'button'; x.addEventListener('click', () => voidForm('op', d, o)); act.append(x); }
      tr.append(el('td', '', ru(o.date)), el('td', '', o.op === 'pay' ? 'Выплата' : 'Взяли / добавили'), el('td', o.op === 'pay' ? 'db-pay' : '', (o.op === 'pay' ? '−' : '+') + money(o.amount, d.currency)),
        el('td', 'db-note', o.note || ''), el('td', 'db-who', (o.created_by || '') + (o.created_at ? ', ' + when(o.created_at) : '')), act);
      tb.append(tr);
    });
    t.append(th, tb); w.append(t); box.append(w);
    const lg = S.log.filter((g) => g.debt_id === d.id);
    if (lg.length) { const det = el('details', 'db-log'); det.append(el('summary', '', 'История изменений (' + lg.length + ')')); lg.forEach((g) => { const r = el('div', 'db-lr'); r.append(el('span', 'db-lw', when(g.at) + ' · ' + (g.who || '')), el('span', '', g.detail)); det.append(r); }); box.append(det); }
    if (!$('db-card').open) $('db-card').showModal();
  }
  // ---------- формы ----------
  const num = (s) => { const v = parseFloat(String(s).replace(/[\s ]/g, '').replace(',', '.')); return isNaN(v) ? null : v; };
  function form(title, fields, submit, onSubmit, hint) {
    const dlg = $('db-form'), f = $('db-form-body'); f.replaceChildren(el('h3', 'pl-modal-title', title)); if (hint) f.append(el('p', 'pl-modal-note', hint));
    const inputs = {};
    fields.forEach((x) => {
      const lab = el('label', 'db-f'); lab.append(el('span', '', x.label + (x.req ? '' : ' (необязательно)')));
      let inp;
      if (x.type === 'select') { inp = el('select'); (x.empty ? [''] : []).concat(x.options).forEach((o) => inp.append(new Option(o || '—', o))); }
      else if (x.type === 'textarea') { inp = el('textarea'); inp.rows = 2; }
      else if (x.type === 'cur') {
        inp = el('div', 'pl-seg db-cur'); inp.setAttribute('role', 'radiogroup'); inp.value = x.value || 'UZS';
        [['UZS', 'Сумы'], ['USD', 'Доллары']].forEach(([v, l]) => { const b = el('button', 'pl-segb', l); b.type = 'button'; b.setAttribute('aria-pressed', v === inp.value ? 'true' : 'false');
          b.addEventListener('click', () => { inp.value = v; inp.querySelectorAll('button').forEach((z) => z.setAttribute('aria-pressed', z === b ? 'true' : 'false')); inp.dispatchEvent(new Event('input')); }); inp.append(b); });
      } else { inp = el('input'); inp.type = x.type === 'date' ? 'date' : 'text'; if (x.type === 'amount') { inp.inputMode = 'decimal'; inp.autocomplete = 'off'; } }
      if (x.type !== 'cur' && x.value != null) inp.value = x.value;
      if (x.max) inp.max = x.max; if (x.req && x.type !== 'cur') inp.required = true; if (x.ph) inp.placeholder = x.ph;
      lab.append(inp);
      if (x.type === 'amount') { const pv = el('small', 'db-pv'); const upd = () => { const v = num(inp.value), cur = x.cur ? x.cur() : 'UZS'; pv.textContent = v ? '= ' + money(v, cur) + (cur === 'USD' ? ' ≈ ' + money(v * S.rate, 'UZS') : '') : ''; }; inp.addEventListener('input', upd); x.watch = upd; lab.append(pv); upd(); }
      inputs[x.name] = inp; f.append(lab);
    });
    fields.forEach((x) => { if (x.type === 'cur') inputs[x.name].addEventListener('input', () => fields.forEach((y) => y.watch && y.watch())); });
    const err = el('p', 'db-err'); err.hidden = true; f.append(err);
    const btns = el('div', 'pl-modal-btns'), cancel = el('button', 'pl-btn', 'Отмена'), ok = el('button', 'pl-btn pl-main', submit);
    cancel.type = 'button'; ok.type = 'submit'; btns.append(cancel, ok); f.append(btns);
    cancel.onclick = () => dlg.close();
    f.onsubmit = async (ev) => {
      ev.preventDefault(); err.hidden = true; ok.disabled = true;
      const val = {}; Object.entries(inputs).forEach(([k, v]) => { val[k] = v.value; });
      try { await onSubmit(val); dlg.close(); await load(); } catch (e) { err.textContent = e.message; err.hidden = false; } finally { ok.disabled = false; }
    };
    dlg.showModal(); const first = f.querySelector('input,select,textarea'); if (first) first.focus();
    return inputs;
  }
  function debtForm(d) {
    const isNew = !d; let ins;
    const F = [
      {name: 'title', label: 'Название долга', req: true, value: d ? d.title : '', ph: 'например: Займ на зарплату, Аренда офиса'},
      {name: 'creditor', label: 'У кого взяли — ФИО или организация', req: true, value: d ? d.creditor : ''},
      {name: 'ctype', label: 'Физлицо или организация', type: 'select', options: S.ctypes, empty: true, value: d ? d.ctype || '' : ''},
      {name: 'kind', label: 'Тип долга', type: 'select', options: S.kinds, req: true, value: d ? d.kind : 'Займ'},
      {name: 'date', label: 'Дата долга', type: 'date', req: true, value: d ? d.opened : S.today, max: S.today}];
    if (isNew) F.push({name: 'currency', label: 'Валюта', type: 'cur', req: true}, {name: 'amount', label: 'Сумма долга', type: 'amount', req: true, cur: () => (ins ? ins.currency.value : 'UZS'), ph: '0'});
    F.push({name: 'due', label: 'Срок возврата', type: 'date', value: d ? d.due || '' : ''}, {name: 'note', label: 'Комментарий', type: 'textarea', value: d ? d.note || '' : '', ph: 'условия, где договорились, номер договора…'});
    ins = form(isNew ? 'Новый долг' : 'Изменить долг', F, isNew ? 'Завести долг' : 'Сохранить',
      (v) => isNew ? api('/api/debts', v) : api('/api/debts/' + d.id, v), isNew ? null : 'Сумму здесь не меняют: если она неверна, аннулируйте операцию в карточке и внесите правильную.');
  }
  function opForm(d, op) {
    form((op === 'pay' ? 'Выплата по долгу' : 'Добавить к долгу') + ' «' + d.title + '»', [
      {name: 'date', label: 'Дата', type: 'date', req: true, value: S.today, max: S.today},
      {name: 'amount', label: 'Сумма, ' + (d.currency === 'USD' ? 'доллары' : 'сумы'), type: 'amount', req: true, cur: () => d.currency, value: op === 'pay' ? '' : '', ph: '0'},
      {name: 'note', label: 'Комментарий', type: 'textarea', ph: op === 'pay' ? 'откуда платили: касса, карта, банк…' : 'за что добавили: аренда за месяц, взяли ещё…'}],
      op === 'pay' ? 'Внести выплату' : 'Добавить', (v) => api('/api/debts/' + d.id + '/ops', Object.assign(v, {op})),
      'Сейчас осталось ' + money(d.left, d.currency) + '.');
  }
  function voidForm(kind, d, o) {
    const t = kind === 'debt' ? 'Аннулировать долг «' + d.title + '»?' : 'Аннулировать операцию от ' + ru(o.date) + ' (' + money(o.amount, d.currency) + ')?';
    form(t, [{name: 'reason', label: 'Причина', type: 'textarea', req: true, ph: kind === 'debt' ? 'например: заведён по ошибке, дубль' : 'например: ошиблись в сумме, внесено дважды'}],
      'Аннулировать', (v) => api(kind === 'debt' ? '/api/debts/' + d.id + '/void' : '/api/debt_ops/' + o.id + '/void', v),
      kind === 'debt' ? 'Долг исчезнет из итогов, но останется в журнале и в Excel на листе «Операции».' : 'Операция останется в истории зачёркнутой, остаток пересчитается.');
  }
  // ---------- события ----------
  $('db-new').addEventListener('click', () => debtForm(null));
  $('db-xlsx').addEventListener('click', () => { location.href = '/api/debts.xlsx'; });
  root.querySelectorAll('[data-dbf]').forEach((b) => b.addEventListener('click', () => { filter = b.dataset.dbf; root.querySelectorAll('[data-dbf]').forEach((x) => x.setAttribute('aria-pressed', x === b ? 'true' : 'false')); render(); }));
  $('db-q').addEventListener('input', (e) => { q = e.target.value.trim().toLowerCase(); render(); });
  $('db-card-close').addEventListener('click', () => $('db-card').close());
  $('db-card').addEventListener('close', () => { cardId = null; });
  ['db-card', 'db-form'].forEach((id) => $(id).addEventListener('click', (ev) => { if (ev.target === ev.currentTarget) ev.currentTarget.close(); }));
  load();
})();
