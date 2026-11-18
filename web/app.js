(() => {
  const DOCS = ['mrr', 'churn', 'clients', 'multi', 'top20', 'revenue', 'cost', 'metrics', 'org', 'calls', 'debts'];
  const CH = {mrr_end: ['stack', 'm'], flow_SFA: ['group', 'm'], flow_POS: ['group', 'm'], retention: ['line', 'p'], survival: ['line', 'p'], churn_n: ['group', 'n'], churn_pct: ['line', 'p'],
    churn_cats: ['hbar', 'p'], seg_clients: ['hbar', 'n'], seg_mrr: ['hbar', 'm'], multi_top: ['hbar', 'm'], top_share: ['hbar', 'p'], top_clients: ['hbar', 'm'], revenue_m: ['stack', 'm'],
    cost_parts: ['hbar', 'm'], hosting: ['stack', 'm'], unit: ['group', 'm'], staff_dep: ['hbar', 'n'], payroll_dep: ['hbar', 'm'], calls_m: ['stack', 'n'], debts: ['group', 'm']};
  const HERO = ['mrr_total', 'subs', 'clients', 'nrr_SFA', 'ltvcac_SFA', 'cost_pct'];
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
  function tile(r, M, prevCut) {
    const t = el('div', 'pl-tile ' + tone(r)), l = el('div', 'pl-tl');
    l.append(el('span', '', r.prod ? r.label.replace(/, (SFA|POS)$/, '') : r.label));
    if (r.prod) l.append(el('span', 'pl-chip ' + r.prod.toLowerCase(), r.prod));
    const row = el('div', 'pl-tr');
    row.append(el('span', 'pl-d ' + tone(r), fmtD(r) + (prevCut && r.delta != null ? ' к ' + prevCut : '')), spark(M.filter((x) => x.key === r.key), r.cut));
    t.append(l, el('div', 'pl-tv', fmt(r.value, r.kind)), row);
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

  function table(e, rows, xs, ss, k) {
    const w = el('div', 'pl-tw'), t = el('table', 'pl-tbl'), hr = el('tr');
    hr.append(el('th', '', ''), ...ss.map((s) => el('th', '', s)));
    const tb = el('tbody');
    xs.forEach((x) => { const tr = el('tr'); tr.append(el('td', '', x), ...ss.map((s) => el('td', '', fmt(rows.get(x + '|' + s), k)))); tb.append(tr); });
    const th = el('thead'); th.append(hr); t.append(th, tb); w.append(t); e.append(w);
  }
  function chart(e, id, S, cut, asTable) {
    const [type, k] = CH[id], data = S.filter((r) => r.cut === cut && r.chart === id);
    e.replaceChildren();
    if (!data.length) { e.append(el('p', 'pl-file', 'Для этого среза данных нет.')); return; }
    const xs = [...new Set(data.map((r) => r.x))], ss = [...new Set(data.map((r) => r.s))], m = new Map(data.map((r) => [r.x + '|' + r.s, r.v]));
    const val = (x, s) => m.get(x + '|' + s) ?? 0;
    if (asTable) return table(e, m, xs, ss, k);
    if (ss.length > 1) { const lg = el('div', 'pl-leg'); ss.forEach((s, i) => { const it = el('span'), sw = el('i'); sw.style.background = color(i); it.append(sw, s); lg.append(it); }); e.append(lg); }
    const W = Math.max(260, e.clientWidth || 600), tip = el('div', 'pl-tip'); tip.hidden = true;
    const showTip = (px, title, lines) => { tip.replaceChildren(el('b', '', title), ...lines.map((t) => el('span', '', t)).flatMap((n) => [n, el('br')])); tip.hidden = false; tip.style.left = Math.max(0, Math.min(px + 12, W - 190)) + 'px'; tip.style.top = '28px'; };
    if (type === 'hbar') {
      const s0 = ss[0], H = Math.max(110, xs.length * 26 + 12), ml = Math.min(W * 0.46, d3.max(xs, (x) => x.length) * 6.6 + 12), mr = 64;
      const svg = d3.create('svg').attr('viewBox', `0 0 ${W} ${H}`).attr('height', H).attr('role', 'img').attr('aria-label', 'Горизонтальные столбцы: ' + xs.slice(0, 4).join(', '));
      const y = d3.scaleBand().domain(xs).range([4, H - 4]).padding(0.25), x = d3.scaleLinear().domain([0, d3.max(xs, (v) => val(v, s0)) || 1]).range([ml, W - mr]);
      const g = svg.selectAll('g').data(xs).join('g');
      g.append('text').attr('class', 'pl-axl').attr('x', ml - 6).attr('y', (d) => y(d) + y.bandwidth() / 2).attr('dy', '0.35em').attr('text-anchor', 'end').text((d) => d.length > ml / 6.6 ? d.slice(0, Math.floor(ml / 6.6) - 1) + '…' : d);
      g.append('rect').attr('x', ml).attr('y', (d) => y(d)).attr('height', y.bandwidth()).attr('width', (d) => Math.max(1, x(val(d, s0)) - ml)).attr('rx', 2).style('fill', color(0));
      g.append('text').attr('class', 'pl-axl').attr('x', (d) => x(val(d, s0)) + 5).attr('y', (d) => y(d) + y.bandwidth() / 2).attr('dy', '0.35em').text((d) => fmt(val(d, s0), k));
      e.append(svg.node()); return;
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
    svg.append('g').selectAll('rect').data(xs).join('rect').attr('x', (x) => band(x)).attr('width', band.step()).attr('y', mt).attr('height', H - mb - mt).attr('fill', 'transparent')
      .on('pointerenter pointermove', (ev, x) => showTip(band(x), x, ss.filter((s) => m.has(x + '|' + s)).map((s) => s + ': ' + fmt(val(x, s), k)).concat(type === 'stack' && ss.length > 1 ? ['Итого: ' + fmt(d3.sum(ss, (s) => val(x, s)), k)] : [])))
      .on('pointerleave', () => { tip.hidden = true; });
    e.append(svg.node(), tip);
  }

  Object.keys(CH).forEach((id) => dash.draw('viz-' + id, (e) => {
    const S = rowsOf('series'), M = rowsOf('metrics'); if (!S || !M) return dash0(e);
    const sec = e.closest('[data-doc]'); chart(e, id, S, curCut(M), sec && view[sec.dataset.doc] === 'table');
  }));
  DOCS.forEach((doc) => {
    dash.draw('kpi-' + doc, (e) => tiles(e, doc));
    dash.draw('file-' + doc, (e) => {
      const M = rowsOf('metrics'); if (!M) return dash0(e);
      const f = M.find((r) => r.cut === curCut(M) && r.doc === doc && r.kind === 'file');
      e.textContent = f ? 'Файл Excel: ' + f.label + ' · собран ' + new Date(f.good).toLocaleString('ru-RU', {dateStyle: 'short', timeStyle: 'short'}) : 'Файл Excel для этого среза не собран.';
    });
  });
  dash.draw('hero', (e) => tiles(e, HERO));
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
  async function refresh(doc) {
    const M = rowsOf('metrics'); if (!M) return;
    let cut = curCut(M);
    if (doc === 'new') {
      const cs = cutsOf(M), l = cs[cs.length - 1], next = new Date(Date.UTC(+l.slice(0, 4), +l.slice(5, 7) + 1, 0)).toISOString().slice(0, 10);
      if (next >= new Date().toISOString().slice(0, 10)) return say('Следующий срез — ' + ru(next) + '. Его можно собрать после окончания месяца; надёжным он станет примерно через три недели.');
      cut = next; doc = 'all';
    }
    say('Запускаю пересборку…');
    try {
      const r = await fetch('/api/refresh', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({cut, docs: doc})});
      const j = await r.json(); if (!r.ok) throw new Error(j.error || 'ошибка сервера');
      say('Пересборка запущена, обычно это меньше минуты.'); if (window.pultPoll) window.pultPoll();
    } catch (err) { say('Не удалось запустить пересборку: ' + err.message); }
  }
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
