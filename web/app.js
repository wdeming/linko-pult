(() => {
  const DOCS = ['mrr', 'churn', 'clients', 'multi', 'top20', 'revenue', 'cash', 'cost', 'metrics', 'org', 'calls', 'debts'];
  // тип графика, формат, опции: row — колонка «Итого» по сериям (и сумма во всплывающей подсказке), col — строка «Итого» по строкам таблицы
  const CH = {
    mrr_end: ['stack', 'm', {row: 'Итого MRR'}], flow_SFA: ['group', 'm', {row: 'Чистое изменение', col: true}], flow_POS: ['group', 'm', {row: 'Чистое изменение', col: true}],
    retention: ['line', 'p', {}], survival: ['line', 'p', {}],
    churn_closed: ['stack', 'n', {row: 'Всего', col: true}], churn_new: ['stack', 'n', {row: 'Всего', col: true}], churn_pct: ['line', 'p', {}],
    churn_cats: ['hbar', 'p', {col: true}], seg_clients: ['hbar', 'n', {col: true}], seg_mrr: ['hbar', 'm', {col: true}], multi_top: ['hbar', 'm', {col: 'Итого по показанным'}],
    top_share: ['hbar', 'p', {}], top_clients: ['hbar', 'm', {col: 'Итого топ-20'}], revenue_m: ['stack', 'm', {row: 'Всего', col: true}],
    cost_parts: ['hbar', 'm', {col: 'Себестоимость всего'}], hosting: ['stack', 'm', {col: true}], unit: ['group', 'm', {}],
    staff_dep: ['hbar', 'n', {col: true}], payroll_dep: ['hbar', 'm', {col: true}], calls_m: ['stack', 'n', {row: 'Всего', col: true}], debts: ['group', 'm', {col: true}],
    cash_prod: ['stack', 'm', {row: 'Поступило всего', col: true}], cash_meth: ['stack', 'm', {row: 'Всего', col: true}], cash_comm: ['stack', 'm', {col: true}]};
  const view = {}; let msgText = '';
  const nf0 = new Intl.NumberFormat('ru-RU', {maximumFractionDigits: 0});
  const nfd = (d) => new Intl.NumberFormat('ru-RU', {minimumFractionDigits: d, maximumFractionDigits: d});
  function fmt(v, k) {
    if (v == null || isNaN(v)) return '—';
    if (k === 'p') return nfd(1).format(v * 100) + '%';
    if (k === 'd1') return nfd(1).format(v);
    if (k === 'd2') return nfd(2).format(v);
    if (k === 'm') { const a = Math.abs(v); if (a >= 1e9) return nfd(2).format(v / 1e9) + ' млрд'; if (a >= 1e6) return nfd(1).format(v / 1e6) + ' млн'; if (a >= 1e3) return nf0.format(v / 1e3) + ' тыс'; }
    return nf0.format(v);
  }
  function fmtD(r) {
    const d = r.delta; if (d == null) return 'нет сравнения';
    if (Math.abs(d) < 1e-9 || (r.kind === 'p' && Math.abs(d) < 0.0005)) return 'без изменений';
    const s = d > 0 ? '+' : '−', a = Math.abs(d);
    if (r.kind === 'p') return s + nfd(1).format(a * 100) + ' п.п.';
    return s + fmt(a, r.kind) + (r.prev ? ' (' + s + nfd(1).format(Math.abs(d / r.prev) * 100) + '%)' : '');
  }
  const tone = (r) => (!r.delta || !r.good || r.kind === 'file' || (r.kind === 'p' && Math.abs(r.delta) < 0.0005)) ? '' : ((r.delta > 0) === (r.good === 'up') ? 'good' : 'bad');
  const rowsOf = (id) => { const g = dash.data(id); return g.status === 'ok' && Array.isArray(g.data) ? g.data : null; };
  const cutsOf = (M) => [...new Set(M.map((r) => r.cut))].sort();
  const curCut = (M) => { const c = dash.params().cut, cs = cutsOf(M); return cs.includes(c) ? c : cs[cs.length - 1]; };
  const ru = (c) => c.slice(8, 10) + '.' + c.slice(5, 7) + '.' + c.slice(0, 4);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const dash0 = (e) => { e.replaceChildren(); e.textContent = '—'; };
  const color = (i) => dash.colors[i % dash.colors.length];

  function spark(rows, cut) {
    const svg = d3.create('svg').attr('class', 'pl-spark').attr('width', 54).attr('height', 20).attr('aria-hidden', 'true');
    const v = rows.filter((r) => r.value != null).sort((a, b) => a.cut < b.cut ? -1 : 1);
    if (v.length < 2) return svg.node();
    const x = d3.scalePoint().domain(v.map((r) => r.cut)).range([3, 51]);
    const ext = d3.extent(v, (r) => r.value), y = d3.scaleLinear().domain(ext[0] === ext[1] ? [ext[0] - 1, ext[1] + 1] : ext).range([17, 3]);
    svg.append('path').attr('d', d3.line().x((r) => x(r.cut)).y((r) => y(r.value))(v)).attr('fill', 'none').style('stroke', 'var(--cds-chart-axis)').attr('stroke-width', 1.5);
    v.forEach((r) => svg.append('circle').attr('cx', x(r.cut)).attr('cy', y(r.value)).attr('r', r.cut === cut ? 2.6 : 1.4).style('fill', r.cut === cut ? 'var(--color-fg)' : 'var(--cds-chart-axis)'));
    return svg.node();
  }
  const tileTip = el('div', 'pl-tip pl-ttip'); tileTip.hidden = true;
  (document.querySelector('.pl') || document.body).append(tileTip);
  function showTileTip(ev, r, M) {
    const rows = M.filter((x) => x.key === r.key).sort((a, b) => a.cut < b.cut ? -1 : 1);
    tileTip.replaceChildren(el('b', '', r.label));
    rows.forEach((x) => { const d = el('div', 'pl-tipr' + (x.cut === r.cut ? ' pl-tiptot' : '')); d.append(el('i'), el('span', '', 'на ' + ru(x.cut)), el('span', 'pl-tipv', fmt(x.value, x.kind))); tileTip.append(d); });
    if (r.delta != null) { const d = el('div', 'pl-tipr'); d.append(el('i'), el('span', '', 'Изменение к прошлому срезу'), el('span', 'pl-tipv pl-d ' + tone(r), fmtD(r))); tileTip.append(d); }
    if (r.good) tileTip.append(el('div', 'pl-tiphint', r.good === 'up' ? 'Рост — хорошо, снижение — плохо' : 'Снижение — хорошо, рост — плохо'));
    tileTip.hidden = false;
    const w = tileTip.offsetWidth, h = tileTip.offsetHeight, vw = document.documentElement.clientWidth, vh = window.innerHeight;
    let x = ev.clientX + 14, y = ev.clientY + 14; if (x + w > vw - 8) x = Math.max(8, ev.clientX - w - 14); if (y + h > vh - 8) y = Math.max(8, ev.clientY - h - 14);
    tileTip.style.left = x + 'px'; tileTip.style.top = y + 'px';
  }
  const hideTileTip = () => { tileTip.hidden = true; };
  window.addEventListener('scroll', hideTileTip, {passive: true});
  function tile(r, M, prevCut) {
    const t = el('div', 'pl-tile ' + tone(r)), l = el('div', 'pl-tl');
    const tot = /, всего$/.test(r.label);
    l.append(el('span', '', r.prod ? r.label.replace(/, (SFA|POS)$/, '') : r.label.replace(/, всего$/, '')));
    if (r.prod) l.append(el('span', 'pl-chip ' + r.prod.toLowerCase(), r.prod));
    else if (tot) l.append(el('span', 'pl-chip all', 'ВСЕГО'));
    const row = el('div', 'pl-tr');
    row.append(el('span', 'pl-d ' + tone(r), fmtD(r) + (prevCut && r.delta != null ? ' к ' + prevCut : '')), spark(M.filter((x) => x.key === r.key), r.cut));
    t.append(l, el('div', 'pl-tv', fmt(r.value, r.kind)), row);
    t.tabIndex = 0;
    t.addEventListener('pointerenter', (ev) => showTileTip(ev, r, M));
    t.addEventListener('pointermove', (ev) => { if (ev.pointerType === 'mouse') showTileTip(ev, r, M); });
    t.addEventListener('pointerleave', hideTileTip);
    t.addEventListener('focus', () => { const b = t.getBoundingClientRect(); showTileTip({clientX: b.left, clientY: b.bottom}, r, M); });
    t.addEventListener('blur', hideTileTip);
    return t;
  }
  function tiles(e, keysOrDoc) {
    const M = rowsOf('metrics'); if (!M) return dash0(e);
    const cut = curCut(M), cs = cutsOf(M), i = cs.indexOf(cut), prev = i > 0 ? ru(cs[i - 1]).slice(0, 5) : '';
    const rows = Array.isArray(keysOrDoc) ? keysOrDoc.map((k) => M.find((r) => r.cut === cut && r.key === k)).filter(Boolean)
      : M.filter((r) => r.cut === cut && r.doc === keysOrDoc && r.kind !== 'file');
    e.replaceChildren(...rows.map((r) => tile(r, M, Array.isArray(keysOrDoc) ? prev : '')));
    if (!rows.length) e.append(el('p', 'pl-file', 'Для этого среза показателей нет.'));
  }

  function table(e, m, xs, ss, k, o) {
    const w = el('div', 'pl-tw'), t = el('table', 'pl-tbl'), hr = el('tr'), v = (x, s) => m.get(x + '|' + s);
    const cols = ss.slice(); if (o.row && ss.length > 1) cols.push(o.row);
    const cell = (x, s) => (s === o.row && ss.length > 1) ? d3.sum(ss, (z) => v(x, z) || 0) : v(x, s);
    hr.append(el('th', '', ''), ...cols.map((s) => el('th', s === o.row ? 'pl-tot' : '', s)));
    const tb = el('tbody');
    xs.forEach((x) => { const tr = el('tr'); tr.append(el('td', '', x), ...cols.map((s) => el('td', s === o.row ? 'pl-tot' : '', fmt(cell(x, s), k)))); tb.append(tr); });
    if (o.col) { const tr = el('tr', 'pl-totrow'); tr.append(el('td', '', o.col === true ? 'Итого' : o.col), ...cols.map((s) => el('td', '', fmt(d3.sum(xs, (x) => cell(x, s) || 0), k)))); tb.append(tr); }
    const th = el('thead'); th.append(hr); t.append(th, tb); w.append(t); e.append(w);
  }
  function chart(e, id, S, cut, asTable) {
    const [type, k, o] = CH[id], data = S.filter((r) => r.cut === cut && r.chart === id);
    e.replaceChildren();
    if (!data.length) { e.append(el('p', 'pl-file', 'Для этого среза данных нет.')); return; }
    const xs = [...new Set(data.map((r) => r.x))], ss = [...new Set(data.map((r) => r.s))], m = new Map(data.map((r) => [r.x + '|' + r.s, r.v]));
    const val = (x, s) => m.get(x + '|' + s) ?? 0;
    if (asTable) return table(e, m, xs, ss, k, o);
    if (ss.length > 1) { const lg = el('div', 'pl-leg'); ss.forEach((s, i) => { const it = el('span'), sw = el('i'); sw.style.background = color(i); it.append(sw, s); lg.append(it); }); e.append(lg); }
    const W = Math.max(260, e.clientWidth || 600), tip = el('div', 'pl-tip'); tip.hidden = true;
    const showTip = (ev, title, lines) => {
      tip.replaceChildren(el('b', '', title)); lines.forEach(([lab, val, i, tot]) => { const r = el('div', tot ? 'pl-tipr pl-tiptot' : 'pl-tipr'), sw = el('i');
        if (i != null) sw.style.background = color(i); r.append(sw, el('span', '', lab), el('span', 'pl-tipv', val)); tip.append(r); });
      tip.hidden = false; const b = e.getBoundingClientRect(), px = ev.clientX - b.left, py = ev.clientY - b.top, tw = tip.offsetWidth, th = tip.offsetHeight;
      tip.style.left = Math.max(0, px + 14 + tw > b.width ? px - tw - 14 : px + 14) + 'px'; tip.style.top = Math.max(0, Math.min(py - th / 2, b.height - th)) + 'px';
    };
    const hide = () => { tip.hidden = true; };
    const linesFor = (x) => { const L = ss.filter((s) => m.has(x + '|' + s)).map((s) => [s, fmt(val(x, s), k), ss.indexOf(s)]);
      if (o.row && ss.length > 1) L.push([o.row, fmt(d3.sum(ss, (s) => val(x, s)), k), null, true]); return L; };
    if (type === 'hbar') {
      const s0 = ss[0], H = Math.max(110, xs.length * 26 + 12), ml = Math.min(W * 0.46, d3.max(xs, (x) => x.length) * 6.6 + 12), mr = 64;
      const svg = d3.create('svg').attr('viewBox', `0 0 ${W} ${H}`).attr('height', H).attr('role', 'img').attr('aria-label', 'Горизонтальные столбцы: ' + xs.slice(0, 4).join(', '));
      const y = d3.scaleBand().domain(xs).range([4, H - 4]).padding(0.25), x = d3.scaleLinear().domain([0, d3.max(xs, (v) => val(v, s0)) || 1]).range([ml, W - mr]);
      const g = svg.selectAll('g').data(xs).join('g');
      g.append('text').attr('class', 'pl-axl').attr('x', ml - 6).attr('y', (d) => y(d) + y.bandwidth() / 2).attr('dy', '0.35em').attr('text-anchor', 'end').text((d) => d.length > ml / 6.6 ? d.slice(0, Math.floor(ml / 6.6) - 1) + '…' : d);
      g.append('rect').attr('x', ml).attr('y', (d) => y(d)).attr('height', y.bandwidth()).attr('width', (d) => Math.max(1, x(val(d, s0)) - ml)).attr('rx', 2).style('fill', color(0));
      g.append('text').attr('class', 'pl-axl').attr('x', (d) => x(val(d, s0)) + 5).attr('y', (d) => y(d) + y.bandwidth() / 2).attr('dy', '0.35em').text((d) => fmt(val(d, s0), k));
      g.append('rect').attr('x', 0).attr('width', W).attr('y', (d) => y(d) - y.step() * y.paddingInner() / 2).attr('height', y.step()).attr('fill', 'transparent').attr('class', 'pl-hit')
        .on('pointerenter pointermove', (ev, d) => { svg.selectAll('.pl-hit').classed('on', (z) => z === d); showTip(ev, d, [[s0, fmt(val(d, s0), k), 0]]); })
        .on('pointerleave', () => { svg.selectAll('.pl-hit').classed('on', false); hide(); });
      e.append(svg.node(), tip); return;
    }
    const H = 250, mt = 8, mb = 24, band = d3.scaleBand().domain(xs).padding(type === 'line' ? 0 : 0.22);
    let lo = 0, hi = 0;
    if (type === 'stack') xs.forEach((x) => { let p = 0, n = 0; ss.forEach((s) => { const v = val(x, s); if (v >= 0) p += v; else n += v; }); hi = Math.max(hi, p); lo = Math.min(lo, n); });
    else { const all = data.map((r) => r.v).filter((v) => v != null); hi = d3.max(all); lo = type === 'line' ? Math.min(d3.min(all), k === 'p' ? d3.min(all) : 0) : Math.min(0, d3.min(all)); if (type === 'line') { const pad = (hi - lo) * 0.12 || 0.05; lo -= pad; hi += pad; } }
    const y = d3.scaleLinear().domain([lo, hi || 1]).nice(4).range([H - mb, mt]);
    const tf = (v) => k === 'p' ? nf0.format(v * 100) + '%' : (k === 'm' ? (Math.abs(v) >= 1e6 ? nf0.format(v / 1e6) + ' млн' : nf0.format(v)) : nf0.format(v));
    const ticks = y.ticks(4), ml = d3.max(ticks, (t) => tf(t).length) * 6.6 + 12;
    band.range([ml, W - 6]);
    const svg = d3.create('svg').attr('viewBox', `0 0 ${W} ${H}`).attr('height', H).attr('role', 'img').attr('aria-label', 'График по ' + xs.length + ' точкам: ' + ss.join(', '));
    const ya = svg.append('g').attr('transform', `translate(${ml},0)`).call(d3.axisLeft(y).tickValues(ticks).tickFormat(tf).tickSize(-(W - 6 - ml)));
    ya.select('.domain').remove();
    const step = Math.max(1, Math.ceil(46 / band.step()));
    svg.append('g').attr('transform', `translate(0,${H - mb})`).call(d3.axisBottom(band).tickSize(0).tickPadding(7).tickValues(xs.filter((_, i) => i % step === 0 || i === xs.length - 1 && (xs.length - 1) % step >= step / 2)));
    if (type === 'line') {
      const cx = (x) => band(x) + band.bandwidth() / 2;
      ss.forEach((s, i) => {
        const pts = xs.filter((x) => m.has(x + '|' + s) && m.get(x + '|' + s) != null);
        svg.append('path').attr('d', d3.line().x((x) => cx(x)).y((x) => y(val(x, s)))(pts)).attr('fill', 'none').attr('stroke-width', 2).style('stroke', color(i));
        svg.append('g').selectAll('circle').data(pts).join('circle').attr('cx', (x) => cx(x)).attr('cy', (x) => y(val(x, s))).attr('r', 2.6).style('fill', color(i));
      });
    } else if (type === 'stack') {
      xs.forEach((x) => { let p = 0, n = 0; ss.forEach((s, i) => { const v = val(x, s); if (!v) return; const y0 = v >= 0 ? p : n, y1 = y0 + v; if (v >= 0) p = y1; else n = y1;
        svg.append('rect').attr('x', band(x)).attr('width', band.bandwidth()).attr('y', y(Math.max(y0, y1))).attr('height', Math.max(0.5, Math.abs(y(y0) - y(y1)))).style('fill', color(i)); }); });
    } else {
      const inner = d3.scaleBand().domain(ss).range([0, band.bandwidth()]).padding(0.08);
      xs.forEach((x) => ss.forEach((s, i) => { const v = val(x, s); svg.append('rect').attr('x', band(x) + inner(s)).attr('width', inner.bandwidth()).attr('y', y(Math.max(0, v))).attr('height', Math.max(0.5, Math.abs(y(v) - y(0)))).style('fill', color(i)); }));
      if (lo < 0) svg.append('line').attr('x1', ml).attr('x2', W - 6).attr('y1', y(0)).attr('y2', y(0)).style('stroke', 'var(--cds-chart-axis)');
    }
    const off = type === 'line' ? 0 : band.step() * band.paddingInner() / 2;
    svg.append('g').selectAll('rect').data(xs).join('rect').attr('class', 'pl-hit').attr('x', (x) => band(x) - off).attr('width', band.step()).attr('y', mt).attr('height', H - mb - mt).attr('fill', 'transparent')
      .on('pointerenter pointermove', (ev, x) => { svg.selectAll('.pl-hit').classed('on', (z) => z === x); showTip(ev, x, linesFor(x)); })
      .on('pointerleave', () => { svg.selectAll('.pl-hit').classed('on', false); hide(); });
    e.append(svg.node(), tip);
  }

  Object.keys(CH).forEach((id) => dash.draw('viz-' + id, (e) => {
    const S = rowsOf('series'), M = rowsOf('metrics'); if (!S || !M) return dash0(e);
    const sec = e.closest('[data-doc]'); chart(e, id, S, curCut(M), !sec || view[sec.dataset.doc] !== 'chart');
  }));
  DOCS.forEach((doc) => {
    dash.draw('kpi-' + doc, (e) => tiles(e, doc));
    dash.draw('file-' + doc, (e) => {
      const M = rowsOf('metrics'); if (!M) return dash0(e);
      const f = M.find((r) => r.cut === curCut(M) && r.doc === doc && r.kind === 'file');
      e.textContent = f ? 'Файл Excel: ' + f.label + ' · собран ' + new Date(f.good).toLocaleString('ru-RU', {dateStyle: 'short', timeStyle: 'short'}) : 'Файл Excel для этого среза не собран.';
    });
  });
  // ---------- блок «Сейчас в Odoo» ----------
  const RMG = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const dayRu = (s) => (+s.slice(8, 10)) + ' ' + RMG[+s.slice(5, 7) - 1];
  const sgn = (v, k) => (v > 0 ? '+' : v < 0 ? '−' : '') + fmt(Math.abs(v), k);
  function ttile(label, value, kind, delta, deltaKind, note, good, chip) {
    const t = el('div', 'pl-tile'), l = el('div', 'pl-tl');
    l.append(el('span', '', label)); if (chip) l.append(el('span', 'pl-chip ' + (chip === 'ВСЕГО' ? 'all' : chip.toLowerCase()), chip));
    let tn = '';
    if (delta && delta.v != null && good) tn = Math.abs(delta.v) < 1 ? '' : ((delta.v > 0) === (good === 'up') ? 'good' : 'bad');
    t.className = 'pl-tile ' + tn;
    t.append(l, el('div', 'pl-tv', typeof value === 'string' ? value : fmt(value, kind)));
    if (delta && delta.v != null) t.append(el('div', 'pl-d ' + tn, sgn(delta.v, deltaKind) + (delta.pct != null && isFinite(delta.pct) ? ' (' + (delta.pct > 0 ? '+' : '−') + nfd(1).format(Math.abs(delta.pct) * 100) + '%)' : '') + ' ' + delta.text));
    if (note) t.append(el('div', 'pl-d', note));
    return t;
  }
  dash.draw('today-sub', (e) => {
    const g = dash.data('today'); if (g.status !== 'ok') { e.textContent = g.status === 'error' ? (g.message || 'Не удалось получить данные') : 'Загружаю данные из Odoo…'; return; }
    const d = g.data; const at = d.fetched_at ? new Date(d.fetched_at).toLocaleString('ru-RU', {dateStyle: 'short', timeStyle: 'short'}) : '';
    e.textContent = 'На ' + dayRu(d.today) + ' ' + d.today.slice(0, 4) + '. Движения, деньги и счета — с 1 ' + RMG[+d.today.slice(5, 7) - 1] + '. Данные получены ' + at + '.' + (d.error ? ' ' + d.error : '');
  });
  dash.draw('today-tiles', (e) => {
    const g = dash.data('today'); if (g.status !== 'ok' || !g.data.mrr) { e.replaceChildren(); return; }
    const d = g.data, P = ['SFA', 'POS'], s = (o) => P.reduce((a, p) => a + (o && o[p] || 0), 0);
    const pe = 'к ' + d.prev_end.slice(8, 10) + '.' + d.prev_end.slice(5, 7), lc = d.last_cut ? 'к срезу ' + d.last_cut.slice(8, 10) + '.' + d.last_cut.slice(5, 7) : '';
    const mv = (et) => s(d.mov[et] || {}), net = mv('0_creation') + mv('1_expansion') + mv('15_contraction') + mv('2_churn');
    const ch = (a, b) => ({v: a - b, pct: b ? a / b - 1 : null});
    const subsNow = s(d.subs);
    const ps = '1–' + (+d.prev_same.slice(8, 10)) + ' ' + RMG[+d.prev_same.slice(5, 7) - 1];
    e.replaceChildren(
      ttile('MRR', s(d.mrr), 'm', Object.assign(ch(s(d.mrr), s(d.mrr_prev)), {text: pe}), 'm', null, 'up', 'ВСЕГО'),
      ttile('MRR', d.mrr.SFA, 'm', Object.assign(ch(d.mrr.SFA, d.mrr_prev.SFA), {text: pe}), 'm', null, 'up', 'SFA'),
      ttile('MRR', d.mrr.POS, 'm', Object.assign(ch(d.mrr.POS, d.mrr_prev.POS), {text: pe}), 'm', null, 'up', 'POS'),
      ttile('Подписок в базе', subsNow, 'n', d.last_subs != null ? Object.assign(ch(subsNow, d.last_subs), {text: lc}) : null, 'n', 'SFA ' + d.subs.SFA + ' · POS ' + d.subs.POS, 'up'),
      ttile('Клиентов', d.clients, 'n', d.last_clients != null ? Object.assign(ch(d.clients, d.last_clients), {text: lc}) : null, 'n', null, 'up'),
      ttile('Изменение MRR с 1-го числа', sgn(net, 'm'), 'm', null, 'm', 'новые ' + sgn(mv('0_creation'), 'm') + ' · расширения ' + sgn(mv('1_expansion'), 'm') + ' · сокращения ' + sgn(mv('15_contraction'), 'm') + ' · отток ' + sgn(mv('2_churn'), 'm'), 'up'),
      ttile('Поступило денег с 1-го числа', d.cash_in, 'm', Object.assign(ch(d.cash_in, d.cash_in_prev), {text: 'к ' + ps}), 'm', 'комиссии Payme / Click / Paynet ' + fmt(d.cash_comm, 'm'), 'up'),
      ttile('Выставлено счетов с 1-го числа', d.inv, 'm', Object.assign(ch(d.inv, d.inv_prev), {text: 'к ' + ps}), 'm', null, 'up'));
  });

  const one = (name, key, f) => dash.draw(name, (e) => { const M = rowsOf('metrics'); if (!M) return dash0(e); const r = M.find((x) => x.cut === curCut(M) && x.key === key); e.textContent = r ? f(r) : '—'; });
  one('h-mrr', 'mrr_total', (r) => fmt(r.value, 'm') + ' сум'); one('h-chg', 'mrr_total', (r) => fmtD(r)); one('h-sfa', 'mrr_sfa', (r) => fmt(r.value, 'm')); one('h-pos', 'mrr_pos', (r) => fmt(r.value, 'm'));
  dash.draw('cutlabel', (e) => { const M = rowsOf('metrics'); e.textContent = M ? ru(curCut(M)) : '—'; });
  dash.draw('cutsel', (sel) => {
    const M = rowsOf('metrics'); if (!M) { sel.replaceChildren(); return; }
    sel.replaceChildren(...cutsOf(M).reverse().map((c) => new Option(ru(c), c))); sel.value = curCut(M);
    sel.onchange = () => dash.setParams({cut: sel.value});
  });
  const last = () => { const s = rowsOf('status'); return s ? s.find((r) => r.id === 'last') : null; };
  dash.draw('status', (e) => {
    const r = last(), when = r && r.at ? new Date(r.at).toLocaleString('ru-RU', {dateStyle: 'short', timeStyle: 'short'}) : '';
    const what = r ? ' Срез ' + (r.cut || '') + ', документы: ' + (r.docs === 'all' ? 'все' : (r.docs || '')) + '.' : '';
    e.textContent = !r ? 'Данные загружены.' : r.state === 'done' ? 'Данные обновлены ' + when + '.' + what : r.state === 'error' ? 'Пересборка не удалась (' + when + '): ' + (r.msg || '') : r.state === 'running' ? 'Идёт пересборка, начата ' + when + '.' + what : 'Запрос на пересборку отправлен ' + when + '.' + what;
  });
  dash.draw('lamp', (e) => { const r = last(); e.className = 'pl-lamp ' + (!r || r.state === 'done' ? 'ok' : r.state === 'error' ? 'bad' : 'run'); });
  dash.draw('msg', (e) => { e.textContent = msgText; });
  const say = (t) => { msgText = t; dash.draw('msg'); };

  function show(doc, link) {
    document.querySelectorAll('.pl-doc').forEach((s) => { s.hidden = s.dataset.doc !== doc; });
    document.querySelectorAll('.pl-tab').forEach((b) => { b.setAttribute('aria-current', b.dataset.tab === doc ? 'true' : 'false'); });
    if (link) dash.setLink(doc);
  }
  document.querySelectorAll('.pl-tab').forEach((b) => b.addEventListener('click', () => show(b.dataset.tab, true)));
  document.querySelectorAll('.pl-segb').forEach((b) => b.addEventListener('click', () => {
    const sec = b.closest('[data-doc]'); if (!sec) return; view[sec.dataset.doc] = b.dataset.view;
    sec.querySelectorAll('.pl-segb').forEach((x) => x.setAttribute('aria-pressed', x === b ? 'true' : 'false'));
    sec.querySelectorAll('[data-draw^="viz-"]').forEach((v) => dash.draw(v.dataset.draw));
  }));
  function download(doc) {
    const M = rowsOf('metrics'); if (!M) return;
    window.location.href = '/api/file?cut=' + curCut(M) + '&doc=' + doc;
  }
  const RMN = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];
  const confirmBox = document.getElementById('new-confirm'), confirmText = document.getElementById('new-confirm-text');
  let pendingNew = null;
  async function post(cut, docs, isNew) {
    say('Запускаю пересборку…');
    try {
      const r = await fetch('/api/refresh', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({cut, docs, new: !!isNew})});
      const j = await r.json(); if (!r.ok) throw new Error(j.error || 'ошибка сервера');
      say(isNew ? 'Собираю срез на ' + ru(cut) + ', обычно это меньше минуты.' : 'Пересборка запущена, обычно это меньше минуты.'); if (window.pultPoll) window.pultPoll();
    } catch (err) { say(err.message); }
  }
  function refresh(doc) {
    const M = rowsOf('metrics'); if (!M) return;
    if (doc !== 'new') { if (confirmBox) confirmBox.hidden = true; return post(curCut(M), doc, false); }
    const cs = cutsOf(M), l = cs[cs.length - 1], ny = +l.slice(0, 4) + (l.slice(5, 7) === '12' ? 1 : 0), nm = l.slice(5, 7) === '12' ? 1 : +l.slice(5, 7) + 1;
    const next = new Date(Date.UTC(ny, nm, 0)).toISOString().slice(0, 10), today = new Date().toISOString().slice(0, 10);
    if (cs.includes(next)) return say('Срез на ' + ru(next) + ' уже собран. Чтобы пересчитать его, выберите срез и нажмите «Обновить всё».');
    if (next >= today) return say('Последний срез — ' + ru(l) + '. Следующий, на ' + ru(next) + ', можно собрать только после окончания месяца, с ' + ru(new Date(Date.UTC(ny, nm, 1)).toISOString().slice(0, 10)) + '.');
    const days = Math.round((Date.parse(today) - Date.parse(next)) / 864e5);
    pendingNew = next;
    if (confirmText) confirmText.textContent = 'Будет собран новый срез: ' + RMN[nm - 1] + ' ' + ny + ' (на ' + ru(next) + '), все документы. ' +
      (days < 21 ? 'Месяц закончился ' + days + ' дн. назад: часть закрытых подписок ещё могут вернуть в работу, поэтому отток может оказаться завышен. Срез можно будет пересчитать позже кнопкой «Обновить всё».' : '');
    if (confirmBox) confirmBox.hidden = false; say('');
  }
  document.getElementById('new-go')?.addEventListener('click', () => { if (confirmBox) confirmBox.hidden = true; if (pendingNew) post(pendingNew, 'all', true); pendingNew = null; });
  document.getElementById('new-cancel')?.addEventListener('click', () => { if (confirmBox) confirmBox.hidden = true; pendingNew = null; say(''); });
  document.querySelectorAll('[data-dl]').forEach((b) => b.addEventListener('click', () => download(b.dataset.dl)));
  document.querySelectorAll('[data-rf]').forEach((b) => b.addEventListener('click', () => refresh(b.dataset.rf)));
  dash.onData(() => {
    const M = rowsOf('metrics'); if (!M) return; const cut = curCut(M);
    document.querySelectorAll('[data-dl]').forEach((b) => { b.disabled = !M.some((r) => r.cut === cut && r.doc === b.dataset.dl && r.kind === 'file'); });
  });
  const h = location.hash.slice(1); show(DOCS.includes(h) ? h : 'mrr', false);
  let open = null;
  window.addEventListener('beforeprint', () => { const cur = document.querySelector('.pl-doc:not([hidden])'); open = cur ? cur.dataset.doc : 'mrr'; document.querySelectorAll('.pl-doc').forEach((s) => { s.hidden = false; }); });
  window.addEventListener('afterprint', () => { if (open) show(open, false); });
})();
