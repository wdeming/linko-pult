# -*- coding: utf-8 -*-
"""Пульт Linko — веб-сервис: вход, данные срезов из своей базы (SQLite), пересборка из Odoo (только чтение), выдача Excel."""
import os, sys, json, sqlite3, subprocess, threading, datetime, re, calendar, base64
from urllib.parse import quote
from flask import Flask, request, session, redirect, jsonify, send_from_directory, Response
from werkzeug.security import check_password_hash
BASE = os.path.dirname(os.path.abspath(__file__)); CALC = os.path.join(BASE, 'calc'); DATA = os.path.join(BASE, 'data'); DB = os.path.join(DATA, 'pult.db')
CFG = json.load(open(os.path.join(DATA, 'config.json'), encoding='utf-8'))
DOCS = {'all', 'mrr', 'churn', 'clients', 'multi', 'top20', 'cost', 'metrics', 'calls', 'org', 'cash', 'exp'}
app = Flask(__name__, static_folder=os.path.join(BASE, 'web'), static_url_path='/static')
app.secret_key = CFG['secret']
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SECURE=True, SESSION_COOKIE_SAMESITE='Lax', PERMANENT_SESSION_LIFETIME=datetime.timedelta(days=30))
LOCK = threading.Lock()
now = lambda: datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
def db():
    c = sqlite3.connect(DB, timeout=30); c.row_factory = sqlite3.Row; return c
def init():
    with db() as c:
        c.executescript("""
        create table if not exists cutdata(cut text primary key, metrics text, series text, measured_at text);
        create table if not exists file(cut text, doc text, name text, at text, body blob, primary key(cut, doc));
        create table if not exists run(id integer primary key autoincrement, started text, finished text, state text, cut text, docs text, msg text, who text);
        create table if not exists snapshot(id integer primary key autoincrement, cut text, measured_at text, metrics text);
        """)
        c.execute("update run set state='error', msg='Прервано перезапуском сервиса', finished=? where state='running'", (now(),))
def ingest(cuts_done):
    out = os.path.join(CALC, 'pult', 'out'); t = now()
    M = json.load(open(os.path.join(out, 'metrics.json'), encoding='utf-8'))['rows']; S = json.load(open(os.path.join(out, 'series.json'), encoding='utf-8'))['rows']
    with db() as c:
        for cut in sorted({r['cut'] for r in M}):
            m = [r for r in M if r['cut'] == cut]; s = [r for r in S if r['cut'] == cut]
            old = c.execute('select measured_at from cutdata where cut=?', (cut,)).fetchone()
            ma = t if (cut in cuts_done or not old) else old['measured_at']
            c.execute('insert or replace into cutdata(cut, metrics, series, measured_at) values(?,?,?,?)', (cut, json.dumps(m, ensure_ascii=False), json.dumps(s, ensure_ascii=False), ma))
            if cut in cuts_done:
                c.execute('insert into snapshot(cut, measured_at, metrics) values(?,?,?)', (cut, t, json.dumps([{'key': r['key'], 'value': r['value']} for r in m if r['kind'] != 'file'], ensure_ascii=False)))
            bp = os.path.join(out, 'bundle_%s.json' % cut)
            if os.path.exists(bp):
                for doc, f in json.load(open(bp, encoding='utf-8'))['files'].items():
                    c.execute('insert or replace into file(cut, doc, name, at, body) values(?,?,?,?,?)', (cut, doc, f['name'], f['at'], base64.b64decode(f['b64'])))
def work(run_id, cut, docs):
    try:
        r = subprocess.run([sys.executable, os.path.join('pult', 'refresh.py'), cut, docs], cwd=CALC, capture_output=True, text=True, timeout=900, env=dict(os.environ, PYTHONIOENCODING='utf-8'))
        if r.returncode != 0 or 'ГОТОВО' not in r.stdout: raise RuntimeError((r.stderr or r.stdout)[-600:])
        ingest(set(cut.split(','))); state, msg = 'done', 'Обновлено'
    except Exception as e:
        state, msg = 'error', re.sub(r'password=\S+', 'password=***', str(e))[-600:]
    with db() as c: c.execute('update run set state=?, msg=?, finished=? where id=?', (state, msg, now(), run_id))
    LOCK.release()
def last_run():
    with db() as c: r = c.execute('select * from run order by id desc limit 1').fetchone()
    return None if not r else {'id': 'last', 'state': r['state'], 'at': r['finished'] or r['started'], 'cut': r['cut'], 'docs': r['docs'], 'msg': r['msg']}
@app.before_request
def guard():
    if request.path in ('/login', '/healthz'): return
    if not session.get('u'):
        return (jsonify(error='Нужен вход'), 401) if request.path.startswith('/api/') else redirect('/login')
LOGIN = """<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex"><title>Пульт Linko — вход</title>
<style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#0f171d;color:#e6edf2;font:15px "Segoe UI",system-ui,sans-serif}form{background:#18232c;border:1px solid #2b3a46;border-radius:10px;padding:28px;width:min(92vw,340px);display:grid;gap:12px}
h1{font-size:22px;margin:0 0 6px;font-weight:600}input{font:inherit;padding:10px;border-radius:8px;border:1px solid #2b3a46;background:#0f171d;color:inherit}button{font:inherit;padding:10px;border:0;border-radius:8px;background:#3ecf8e;color:#0f171d;font-weight:600;cursor:pointer}p{margin:0;color:#ff7b6b;font-size:13px}</style></head>
<body><form method="post"><h1>Пульт Linko</h1><input name="login" placeholder="Логин" autocomplete="username" required><input name="password" type="password" placeholder="Пароль" autocomplete="current-password" required><button>Войти</button>%s</form></body></html>"""
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u = request.form.get('login', '').strip().lower(); h = CFG['users'].get(u)
        if h and check_password_hash(h, request.form.get('password', '')):
            session.permanent = True; session['u'] = u; return redirect('/')
        return LOGIN % '<p>Неверный логин или пароль.</p>', 401
    return LOGIN % ''
@app.route('/logout')
def logout():
    session.clear(); return redirect('/login')
@app.route('/healthz')
def healthz(): return 'ok'
@app.route('/')
def index():
    r = send_from_directory(app.static_folder, 'index.html'); r.headers['Cache-Control'] = 'no-store'; return r
@app.after_request
def nocache(r):
    r.headers.setdefault('Cache-Control', 'no-store'); return r
@app.route('/api/data')
def api_data():
    with db() as c: rows = c.execute('select * from cutdata order by cut').fetchall()
    M, S = [], []
    for r in rows: M += json.loads(r['metrics']); S += json.loads(r['series'])
    return jsonify(metrics=M, series=S, status=last_run(), measured={r['cut']: r['measured_at'] for r in rows})
@app.route('/api/status')
def api_status(): return jsonify(status=last_run())
@app.route('/api/file')
def api_file():
    with db() as c: f = c.execute('select name, body from file where cut=? and doc=?', (request.args.get('cut', ''), request.args.get('doc', ''))).fetchone()
    if not f: return 'Файл для этого среза не собран', 404
    return Response(bytes(f['body']), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': 'attachment; filename="linko_%s_%s.xlsx"; filename*=UTF-8\'\'%s' % (request.args.get('doc', ''), request.args.get('cut', ''), quote(f['name']))})
@app.route('/api/refresh', methods=['POST'])
def api_refresh():
    j = request.get_json(silent=True) or {}; cut = str(j.get('cut', '')); docs = str(j.get('docs', 'all'))
    m = re.fullmatch(r'2026-(0[89]|1[0-2])-(\d\d)', cut)
    if not m or int(m.group(2)) != calendar.monthrange(2026, int(m.group(1)))[1]: return jsonify(error='Срез должен быть концом месяца 2026 года, начиная с августа'), 400
    if cut >= datetime.date.today().isoformat(): return jsonify(error='Этот месяц ещё не закончился'), 400
    if not docs or any(d not in DOCS for d in docs.split(',')): return jsonify(error='Неизвестный документ'), 400
    with db() as c:
        have = [r[0] for r in c.execute('select cut from cutdata order by cut')]
    if j.get('new'):
        if cut in have: return jsonify(error='Срез на %s уже собран. Чтобы пересчитать его, выберите срез и нажмите «Обновить всё».' % cut), 409
        if have:
            ly, lm = int(have[-1][:4]), int(have[-1][5:7]); ny, nm = (ly + 1, 1) if lm == 12 else (ly, lm + 1)
            if cut != '%04d-%02d-%02d' % (ny, nm, calendar.monthrange(ny, nm)[1]): return jsonify(error='Новым может быть только следующий месяц после последнего среза (%s).' % have[-1]), 400
    elif cut not in have:
        return jsonify(error='Такого среза ещё нет. Новый месяц собирается кнопкой «Собрать новый месяц».'), 400
    if not LOCK.acquire(blocking=False): return jsonify(error='Пересборка уже идёт, дождитесь окончания'), 409
    with db() as c: rid = c.execute("insert into run(started, state, cut, docs, msg, who) values(?,?,?,?,?,?)", (now(), 'running', cut, docs, 'Пересборка запущена', session.get('u'))).lastrowid
    threading.Thread(target=work, args=(rid, cut, docs), daemon=True).start()
    return jsonify(ok=True, run=rid)
TODAY = {'at': 0, 'data': None, 'err': None}
TLOCK = threading.Lock()
@app.route('/api/today')
def api_today():
    import time
    force = request.args.get('force') == '1'
    with TLOCK:
        if force or not TODAY['data'] or time.time() - TODAY['at'] > 600:
            try:
                r = subprocess.run([sys.executable, os.path.join('pult', 'today.py')], cwd=CALC, capture_output=True, text=True, timeout=90, env=dict(os.environ, PYTHONIOENCODING='utf-8'))
                if r.returncode != 0: raise RuntimeError((r.stderr or r.stdout)[-300:])
                TODAY.update(at=time.time(), data=json.loads(r.stdout), err=None)
            except Exception as e:
                TODAY['err'] = 'Не удалось получить данные из Odoo: ' + re.sub(r'password=\S+', 'password=***', str(e))[-200:]
        d = dict(TODAY['data'] or {})
    with db() as c:
        row = c.execute('select cut, metrics from cutdata order by cut desc limit 1').fetchone()
    if row:
        last = {m['key']: m['value'] for m in json.loads(row['metrics'])}
        d['last_cut'] = row['cut']; d['last_subs'] = last.get('subs'); d['last_clients'] = last.get('clients')
    d['fetched_at'] = datetime.datetime.fromtimestamp(TODAY['at'], datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ') if TODAY['at'] else None
    d['error'] = TODAY['err']
    return jsonify(d)
@app.route('/api/history')
def api_history():
    with db() as c: rows = c.execute('select id, started, finished, state, cut, docs, msg, who from run order by id desc limit 50').fetchall()
    return jsonify(runs=[dict(r) for r in rows])
init()
