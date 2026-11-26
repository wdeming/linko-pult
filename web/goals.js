/* Раздел «Цели и планы» → «Планы продаж» (26.11.2026). Данные — /api/goals/sales.
   План — новый MRR (сумы) по месяцу и продукту; факт — новый MRR + расширения/сокращения подписок младше 6 месяцев. */
(function () {
  const $ = (id) => document.getElementById(id);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const MON = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const MONG = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
  const KIND = {new: 'Новая подписка', exp: 'Расширение', con: 'Сокращение', churn: 'Отток (справочно, не вычитается)'};
  const nf = (d) => new Intl.NumberFormat('ru-RU', {minimumFractionDigits: d, maximumFractionDigits: d});
  const mln = (v) => v == null ? '—' : (Math.abs(v) >= 1e6 ? nf(1).format(v / 1e6) + ' млн' : Math.abs(v) >= 1e3 ? nf(0).format(v / 1e3) + ' тыс' : nf(0).format(v));
  const sum = (v) => nf(0).format(Math.round(v));
  const pct = (a, b) => b ? Math.round(a / b * 100) : null;
  const ymLab = (ym) => MON[+ym.slice(5) - 1] + ' ' + ym.slice(0, 4);
  let D = null, year = null, sel = null, busy = false;
  const planOf = (ym, p) => { const r = D.plans.find((x) => x.ym === ym && x.prod === p); return r ? r.amount : null; };
  const factOf = (ym, p) => { const f = (D.fact[ym] || {})[p]; return f ? f.total : null; };
  const planTot = (ym) => { const v = D.prods.map((p) => planOf(ym, p)).filter((x) => x != null); return v.length ? v.reduce((a, b) => a + b, 0) : null; };
  const factTot = (ym) => { const v = D.prods.map((p) => factOf(ym, p)).filter((x) => x != null); return v.length ? v.reduce((a, b) => a + b, 0) : (ym >= '2026-02' && ym <= D.current ? 0 : null); };
  async function load(force) {
    if (busy) return; busy = true; const sub = $('sp-sub'); if (sub) sub.textContent = force ? 'Пересчитываю факт продаж из учётной системы…' : 'Загружаю…';
    try {
      const r = await fetch('/api/goals/sales' + (force ? '?force=1' : '')); if (r.status === 401) { location.href = '/login'; return; }
      const j = await r.json(); if (!r.ok) throw new Error(j.error || 'Не удалось загрузить планы'); D = j; if (!year) year = D.current.slice(0, 4); draw();
    } catch (e) { if (sub) sub.textContent = e.message; } finally { busy = false; }
  }
  window.loadSales = (force) => load(force);
  function cell(f, p, future) {
    const td = el('td', 'sp-c');
    if (f == null && p == null) { td.textContent = '—'; td.classList.add('sp-none'); return td; }
    if (future || f == null) { td.append(el('div', 'sp-plan', p != null ? 'план ' + mln(p) : '—')); return td; }
    td.append(el('div', 'sp-f', mln(f)));
    if (p != null) {
      const k = pct(f, p), bar = el('div', 'sp-bar'), fill = el('i'); fill.style.width = Math.max(0, Math.min(100, k || 0)) + '%';
      fill.className = k >= 100 ? 'ok' : k >= 70 ? 'mid' : 'bad'; bar.append(fill);
      td.append(el('div', 'sp-plan', 'из ' + mln(p) + ' · ' + (k == null ? '—' : k + '%')), bar);
    } else td.append(el('div', 'sp-plan', 'план не задан'));
    return td;
  }
  function tiles() {
    const cur = D.current, now = new Date(), dim = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate(), share = now.getDate() / dim;
    const pc = planTot(cur), fc = factTot(cur) || 0, y0 = cur.slice(0, 4) + '-01';
    const ms = []; for (let m = y0; m <= cur; m = next(m)) ms.push(m);
    const fy = ms.reduce((a, m) => a + (factTot(m) || 0), 0), pm = ms.filter((m) => planTot(m) != null), py = pm.reduce((a, m) => a + planTot(m), 0), fpy = pm.reduce((a, m) => a + (factTot(m) || 0), 0);
    const t = (lab, val, note, cls) => { const x = el('div', 'pl-tile ' + (cls || '')); const l = el('div', 'pl-tl'); l.append(el('span', '', lab)); x.append(l, el('div', 'pl-tv', val)); if (note) x.append(el('div', 'pl-d', note)); return x; };
    const k = pct(fc, pc), fcst = share > 0 ? fc / share : fc, ky = pct(fpy, py);
    $('sp-tiles').replaceChildren(
      t('План на ' + MONG[+cur.slice(5) - 1], pc == null ? 'не задан' : mln(pc), 'новый MRR, все продукты'),
      t('Факт с 1-го числа', mln(fc), k == null ? 'план не задан' : 'исполнение ' + k + '%', k == null ? '' : k >= Math.round(share * 100) ? 'good' : 'bad'),
      t('Прогноз на конец месяца', mln(fcst), pc ? 'при нынешнем темпе — ' + pct(fcst, pc) + '% плана' : 'при нынешнем темпе'),
      t('С начала года', mln(fy), py ? 'месяцы с планом: ' + mln(fpy) + ' из ' + mln(py) + ' · ' + ky + '%' : 'план за год не задан', py ? (ky >= 100 ? 'good' : 'bad') : ''));
  }
  function next(ym) { let y = +ym.slice(0, 4), m = +ym.slice(5) + 1; if (m > 12) { m = 1; y++; } return y + '-' + String(m).padStart(2, '0'); }
  function table() {
    const t = $('sp-tbl'), hr = el('tr'); hr.append(el('th', '', 'Месяц'), ...D.prods.map((p) => el('th', '', D.pname[p])), el('th', 'pl-tot', 'Итого'));
    const tb = el('tbody');
    for (let m = 1; m <= 12; m++) {
      const ym = year + '-' + String(m).padStart(2, '0'), fut = ym > D.current, tr = el('tr', 'sp-row' + (ym === sel ? ' sp-sel' : '') + (ym === D.current ? ' sp-cur' : ''));
      tr.append(el('td', 'sp-m', MON[m - 1] + (ym === D.current ? ' · идёт' : '')));
      D.prods.forEach((p) => tr.append(cell(fut ? null : factOf(ym, p) ?? (ym >= '2026-02' ? 0 : null), planOf(ym, p), fut)));
      const tc = cell(fut ? null : factTot(ym), planTot(ym), fut); tc.classList.add('pl-tot'); tr.append(tc);
      if (!fut && ym >= '2026-02') tr.addEventListener('click', () => { sel = ym; draw(); $('sp-det').scrollIntoView({behavior: 'smooth', block: 'start'}); });
      tb.append(tr);
    }
    const th = el('thead'); th.append(hr); t.replaceChildren(th, tb);
  }
  function details() {
    const box = $('sp-det'); if (!sel) { box.replaceChildren(); return; }
    const rows = D.details[sel] || [], h = el('h3', 'sp-h', 'Из чего сложился факт: ' + MONG[+sel.slice(5) - 1] + ' ' + sel.slice(0, 4));
    if (!rows.length) { box.replaceChildren(h, el('p', 'pl-file', 'Новых продаж и изменений новых подписок в этом месяце нет.')); return; }
    const w = el('div', 'pl-tw'), t = el('table', 'pl-tbl sp-dt'), hr = el('tr');
    ['Продукт', 'Подписка', 'Клиент', 'Что', 'MRR', 'Дата', 'Создана', 'Менеджер'].forEach((x) => hr.append(el('th', '', x)));
    const tb = el('tbody');
    D.prods.forEach((p) => {
      const ORD = {new: 0, exp: 1, con: 2, churn: 3};
      const rs = rows.filter((r) => r.prod === p).sort((a, b) => ORD[a.kind] - ORD[b.kind] || Math.abs(b.a) - Math.abs(a.a)); if (!rs.length) return;
      const by = (k) => rs.filter((r) => r.kind === k).reduce((a, r) => a + r.a, 0), tot = by('new') + by('exp') + by('con'), g = el('tr', 'sp-grp');
      const c1 = el('td', '', D.pname[p] + ' — продажи ' + mln(tot) + ': новые ' + mln(by('new')) + ', расширения ' + mln(by('exp')) + ', сокращения ' + mln(by('con')) + (by('churn') ? ' · справочно: отток новых подписок ' + mln(by('churn')) : ''));
      c1.colSpan = 8; g.append(c1); tb.append(g);
      rs.forEach((r) => { const tr = el('tr', r.kind === 'churn' ? 'sp-info' : ''); [D.pname[p], r.name, r.client, KIND[r.kind], sum(r.a), r.date.split('-').reverse().join('.'), r.born.split('-').reverse().join('.'), r.user].forEach((x, i) => tr.append(el('td', i === 4 ? (r.a < 0 ? 'db-bad' : '') : '', x))); tb.append(tr); });
    });
    const th = el('thead'); th.append(hr); t.append(th, tb); w.append(t); box.replaceChildren(h, w);
  }
  function form() {
    const f = $('sp-form'); f.hidden = !D.can_edit; if (!D.can_edit) return;
    const ys = $('sp-fy'), ms = $('sp-fm'), ps = $('sp-fp');
    if (!ys.options.length) {
      for (let y = +D.first.slice(0, 4); y <= +D.last.slice(0, 4); y++) ys.append(new Option(y, y));
      MON.forEach((m, i) => ms.append(new Option(m, String(i + 1).padStart(2, '0'))));
      D.prods.forEach((p) => ps.append(new Option(D.pname[p], p)));
      ys.value = D.current.slice(0, 4); ms.value = D.current.slice(5);
      [ys, ms, ps].forEach((x) => x.addEventListener('change', fill));
      $('sp-fa').addEventListener('input', (e) => { const d = e.target.value.replace(/\D/g, ''); e.target.value = d ? nf(0).format(+d) : ''; });
      f.addEventListener('submit', (e) => { e.preventDefault(); save(false); });
      $('sp-del').addEventListener('click', () => save(true));
    }
    fill();
  }
  function fill() { const ym = $('sp-fy').value + '-' + $('sp-fm').value, v = planOf(ym, $('sp-fp').value); $('sp-fa').value = v == null ? '' : nf(0).format(v); $('sp-msg').textContent = ym < D.first || ym > D.last ? 'Вне горизонта: план можно задать с ' + ymLab(D.first) + ' по ' + ymLab(D.last) : ''; }
  async function save(del) {
    const ym = $('sp-fy').value + '-' + $('sp-fm').value, prod = $('sp-fp').value, raw = $('sp-fa').value.replace(/\D/g, '');
    if (!del && !raw) { $('sp-msg').textContent = 'Укажите сумму в сумах'; return; }
    const r = await fetch('/api/goals/sales', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ym, prod, amount: del ? null : +raw})});
    const j = await r.json(); if (!r.ok) { $('sp-msg').textContent = j.error || 'Не удалось сохранить'; return; }
    $('sp-msg').textContent = (del ? 'План удалён: ' : 'Сохранено: ') + D.pname[prod] + ', ' + ymLab(ym) + (del ? '' : ' — ' + nf(0).format(+raw) + ' сум');
    year = ym.slice(0, 4); await load(false);
  }
  function draw() {
    if (!D) return;
    const at = D.computed ? new Date(D.computed).toLocaleString('ru-RU', {dateStyle: 'short', timeStyle: 'short'}) : '—';
    $('sp-sub').textContent = 'Факт — из учётной системы, получен ' + at + (D.error ? '. ' + D.error : '') + '. Нажмите на месяц, чтобы увидеть, из каких подписок сложился факт.';
    const yy = $('sp-year'); if (!yy.options.length) { for (let y = 2026; y <= +D.last.slice(0, 4); y++) yy.append(new Option(y + ' год', y)); yy.addEventListener('change', () => { year = yy.value; sel = null; draw(); }); }
    yy.value = year; $('sp-refresh').hidden = !D.can_edit;
    form(); tiles(); table(); details();
    const lg = $('sp-log'); lg.replaceChildren(...(D.log.length ? D.log.slice(0, 30).map((x) => el('li', '', x.at.slice(0, 16).replace('T', ' ') + ' · ' + x.who + ' · ' + D.pname[x.prod] + ', ' + ymLab(x.ym) + ': ' + (x.old == null ? '—' : sum(x.old)) + ' → ' + (x.new == null ? 'удалён' : sum(x.new)))) : [el('li', '', 'Изменений ещё не было')]));
  }
  const init = () => { const b = $('sp-refresh'); if (b && !b.dataset.on) { b.dataset.on = 1; b.addEventListener('click', () => load(true)); } const s = $('goal_sales'); if (s && !s.hidden && !D) load(false); };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
  setTimeout(init, 300);   // раздел мог открыться по ссылке раньше, чем загрузился этот файл
})();
