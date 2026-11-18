// Мини-среда для страницы пульта: данные, рисовальщики, параметры
window.dash = (() => {
  const D = {}, draws = {}, ons = []; const params = {cut: ''};
  const visible = (el) => !el.closest('[hidden]');
  const run = () => { ons.forEach((f) => { try { f(); } catch (e) { console.error(e); } });
    document.querySelectorAll('[data-draw]').forEach((el) => { const f = draws[el.dataset.draw]; if (f && visible(el)) { try { f(el); } catch (e) { console.error(el.dataset.draw, e); } } }); };
  return {
    colors: ['var(--c1)', 'var(--c2)', 'var(--c3)', 'var(--c4)', 'var(--c5)', 'var(--c6)'],
    data: (id) => D[id] || {status: 'loading', data: []},
    onData: (f) => { ons.push(f); },
    draw: (name, f) => { if (f) { draws[name] = f; return; } document.querySelectorAll('[data-draw="' + name + '"]').forEach((el) => { if (draws[name] && visible(el)) draws[name](el); }); },
    params: () => params, setParams: (p) => { Object.assign(params, p); run(); }, resetParams: () => { params.cut = ''; run(); },
    setLink: (n) => { history.replaceState(null, '', '#' + n); },
    _set: (id, rows) => { D[id] = {status: 'ok', data: rows}; }, _run: run,
  };
})();
