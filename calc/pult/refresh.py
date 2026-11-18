# -*- coding: utf-8 -*-
"""Пересборка данных пульта.  python3 pult/refresh.py 2026-10-31[,2026-09-30] [all|mrr,churn,clients,multi,top20,cost,metrics,calls,org]
Считает срез(ы) в pult/work/<срез>/ и собирает pult/out/ (metrics.json, series.json, bundle_<срез>.json)."""
import os, sys, subprocess, json, datetime
P = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(P)
cuts = sys.argv[1].split(','); docs = (sys.argv[2] if len(sys.argv) > 2 else 'all').split(',')
ALL = ['mrr', 'churn', 'clients', 'multi', 'top20', 'cost', 'metrics', 'calls', 'org']
if 'all' in docs: docs = ALL
if 'multi' in docs and 'clients' not in docs: docs.append('clients')
if 'metrics' in docs and 'cost' not in docs: docs.append('cost')
RUN = "import sys,runpy; sys.path.insert(0,%r); sys.path.insert(1,%r); runpy.run_path(sys.argv[1], run_name='__main__')" % (P, ROOT)
def run(script, env):
    r = subprocess.run([sys.executable, '-c', RUN, script], cwd=ROOT, env=env, capture_output=True, text=True, timeout=170)
    if r.returncode: raise SystemExit('ОШИБКА %s:\n%s' % (script, (r.stderr or r.stdout)[-1500:]))
log = []
for cut in cuts:
    out = os.path.join(P, 'work', cut); os.makedirs(out, exist_ok=True)
    env = dict(os.environ, LINKO_CUT=cut, LINKO_OUT=out, PYTHONIOENCODING='utf-8')
    for s in ('extract_oct31.py', 'extract_oct31_mkt.py', 'extract_oct31_churn.py', 'extract_oct31_calls.py', os.path.join(P, 'extract_staff.py')): run(s, env)
    for d in ('clients', 'multi', 'top20', 'cost', 'metrics', 'mrr', 'churn', 'calls'):
        if d in docs: run('gen_oct31_%s.py' % d, env)
    st = os.path.join(out, 'stamp.json'); old = json.load(open(st)) if os.path.exists(st) else {}
    now = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    for d in docs: old[d] = now
    json.dump(old, open(st, 'w')); log.append('%s: %s' % (cut, ','.join(docs)))
run(os.path.join(P, 'build_data.py'), dict(os.environ, PYTHONIOENCODING='utf-8'))
print('ГОТОВО', '; '.join(log))
