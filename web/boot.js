(() => {
  let timer = null, lastState = null;
  async function load() {
    const r = await fetch('/api/data'); if (r.status === 401) { location.href = '/login'; return; }
    const j = await r.json(); dash._set('metrics', j.metrics); dash._set('series', j.series); dash._set('status', j.status ? [j.status] : []); dash._set('bundles', []);
    lastState = j.status ? j.status.state : null; dash._run();
    if (lastState === 'running') poll();
  }
  async function tick() {
    try { const r = await fetch('/api/status'); const j = await r.json(); dash._set('status', j.status ? [j.status] : []); dash.draw('status'); dash.draw('lamp');
      if (j.status && j.status.state !== 'running') { clearInterval(timer); timer = null; if (lastState === 'running') await load(); }
      lastState = j.status ? j.status.state : null; } catch (e) { /* сеть моргнула — попробуем ещё раз */ }
  }
  function poll() { lastState = 'running'; if (!timer) timer = setInterval(tick, 3000); tick(); }
  window.pultPoll = poll;
  new MutationObserver((ms) => { if (ms.some((m) => m.target.classList && m.target.classList.contains('pl-doc'))) dash._run(); }).observe(document.getElementById('dash-root'), {attributes: true, attributeFilter: ['hidden'], subtree: true});
  let rz = null; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => dash._run(), 150); });
  load().catch((e) => console.error(e));
})();
