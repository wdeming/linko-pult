# -*- coding: utf-8 -*-
"""Пульт Linko — веб-сервис: вход, данные срезов из своей базы (SQLite), пересборка из Odoo (только чтение), выдача Excel."""
import os, sys, json, sqlite3, subprocess, threading, datetime, re, calendar, base64
from uchet import walk as uchet_walk, xlsx as uchet_xlsx   # «Odoo» → «учётная система» во всём, что видит пользователь
from urllib.parse import quote
from flask import Flask, request, session, redirect, jsonify, send_from_directory, Response
from werkzeug.security import check_password_hash
BASE = os.path.dirname(os.path.abspath(__file__)); CALC = os.path.join(BASE, 'calc'); DATA = os.path.join(BASE, 'data'); DB = os.path.join(DATA, 'pult.db')
CFG = json.load(open(os.path.join(DATA, 'config.json'), encoding='utf-8'))
DOCS = {'all', 'mrr', 'churn', 'clients', 'multi', 'top20', 'cost', 'metrics', 'calls', 'org', 'cash', 'exp', 'fin'}
app = Flask(__name__, static_folder=os.path.join(BASE, 'web'), static_url_path='/static')
app.secret_key = CFG['secret']
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SECURE=True, SESSION_COOKIE_SAMESITE='Lax', PERMANENT_SESSION_LIFETIME=datetime.timedelta(days=30))
LOCK = threading.Lock()
def role():
    """Роль пользователя: admin — всё; viewer — только просмотр (без пересборок, смены среза и ввода данных). data/config.json → perms.<логин>.role"""
    return ((CFG.get('perms') or {}).get(session.get('u') or '', {}) or {}).get('role', 'admin')
RO = ('viewer', 'bank')     # роли только для просмотра
# ---- доступ для банка (роль bank, 26.11.2026): что убрано из данных и файлов ----
BANK_DOCS = {'debts', 'fin_cash', 'fin_bal', 'fin_src', 'fin_plan', 'exp'}                     # вкладки целиком
BANK_CHARTS = {'mrr_jerr', 'payroll_dep', 'ar_age', 'cash_tab', 'cash_flow', 'cash_bal', 'cash_acc', 'bal_tab', 'src_tab', 'src_check', 'exp_cash', 'exp_period', 'exp_flow', 'exp_cat'}
BANK_KEYS = {'f_ar', 'f_ar90', 'f_unm', 'h_capex_tap', 'cash_cum'}
BANK_ROWS = {('hist_tab', 'Справочно: EBITDA в исходном файле'), ('seg_mrr', 'Расхождения журнала Odoo')}
BANK_FILES_BLOCK = {'debts', 'org', 'calls', 'exp'}                                              # файлы с персональными данными и внутренним учётом
BANK_SHEETS_DROP = {'fin': ('Деньги и запас', 'Баланс на срез', 'Источники и методика'), 'cost': ('Сотрудники отделов',)}
def bank_rows(M, S):
    M = [r for r in M if r['doc'] not in BANK_DOCS and r['key'] not in BANK_KEYS]
    S = [r for r in S if r['chart'] not in BANK_CHARTS and (r['chart'], r['x']) not in BANK_ROWS]
    return M, S
def bank_xlsx(doc, body):
    import io, openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(body))
    for nm in BANK_SHEETS_DROP.get(doc, ()):
        if nm in wb.sheetnames: del wb[nm]
    for ws in wb.worksheets:      # строки, которые банку не показываем
        for r in range(ws.max_row, 0, -1):
            a = ws.cell(r, 1).value
            if isinstance(a, str) and (a.startswith('Справочно: EBITDA в исходном файле') or a.startswith('Дебиторка на') or a.startswith('ИТОГО неоплаченных') or a.startswith('Поступления, не привязанные')
                                       or a in ('не просрочено', 'до 30 дней', '31–90 дней', '91–180 дней', 'больше 180 дней') or a.startswith('Дебиторка завышена') or 'сторнировать' in a or '201000' in a):
                ws.delete_rows(r)
    out = io.BytesIO(); wb.save(out); return out.getvalue()
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
    if request.path.startswith('/api/goals') and role() == 'bank': return jsonify(error='Нет доступа'), 403   # раздел «Цели и планы» банку не показывается
LOGIN = """<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex"><title>Пульт Linko — вход</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Manrope:wght@700;800&display=swap">
<style>:root{--a:#3b5bdb;--b:#0ca678}*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;padding:16px;font:15px Inter,"Segoe UI",system-ui,sans-serif;color:#141b2d;
background:radial-gradient(900px 500px at 15%% 0%%,#dbe4ff,transparent 60%%),radial-gradient(800px 500px at 100%% 100%%,#c3fae8,transparent 60%%),#f3f5fa}
form{background:#fff;border:1px solid #e2e7f0;border-radius:18px;padding:30px;width:min(100%%,360px);display:grid;gap:12px;box-shadow:0 2px 4px rgba(20,27,45,.05),0 20px 50px rgba(20,27,45,.12)}
.m{width:46px;height:46px;border-radius:13px;display:grid;place-items:center;color:#fff;font:800 24px Manrope,sans-serif;background:linear-gradient(135deg,var(--a),var(--b));box-shadow:0 8px 20px rgba(59,91,219,.35)}
h1{font:800 24px Manrope,sans-serif;margin:6px 0 0;letter-spacing:-.01em}.s{margin:0 0 8px;color:#5f6b80;font-size:13px}
input{font:inherit;padding:11px 12px;border-radius:11px;border:1px solid #d5dce8;background:#f8faff;color:inherit;outline:none}input:focus{border-color:var(--a);box-shadow:0 0 0 3px rgba(59,91,219,.15)}
button{font:600 15px Inter,sans-serif;padding:11px;border:0;border-radius:11px;background:linear-gradient(135deg,var(--a),#4c6ef5);color:#fff;cursor:pointer;box-shadow:0 6px 16px rgba(59,91,219,.3)}
p.e{margin:0;color:#d6336c;font-size:13px}
@media (prefers-color-scheme:dark){body{color:#e7ecf6;background:radial-gradient(900px 500px at 15%% 0%%,#1c2a55,transparent 60%%),radial-gradient(800px 500px at 100%% 100%%,#0f3a33,transparent 60%%),#0b1220}
form{background:#131c2e;border-color:#24314b}input{background:#0f1726;border-color:#24314b}.s{color:#93a1ba}}</style></head>
<body><form method="post"><div class="m">L</div><h1>Пульт Linko</h1><p class="s">Данные учётной системы, срезы и финансы</p><input name="login" placeholder="Логин" autocomplete="username" required><input name="password" type="password" placeholder="Пароль" autocomplete="current-password" required><button>Войти</button>%s</form></body></html>"""
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u = request.form.get('login', '').strip().lower(); h = CFG['users'].get(u)
        if h and check_password_hash(h, request.form.get('password', '')):
            session.permanent = True; session['u'] = u; return redirect('/')
        return LOGIN % '<p class="e">Неверный логин или пароль.</p>', 401
    return LOGIN % ''
@app.route('/logout')
def logout():
    session.clear(); return redirect('/login')
@app.route('/healthz')
def healthz(): return 'ok'
@app.route('/')
def index():
    html = open(os.path.join(app.static_folder, 'index.html'), encoding='utf-8').read()
    if role() == 'viewer': html = html.replace('<body>', '<body class="pl-ro">', 1)
    if role() == 'bank':
        html = html.replace('<body>', '<body class="pl-ro pl-bank">', 1)
        for t in (' Дебиторка завышена: много поступлений в учётной системе не привязано к счетам, поэтому оплаченные счета числятся неоплаченными.',
                  ', поэтому EBITDA 2025 лучше, чем в файле финансиста, ровно на эти суммы (строка «EBITDA в исходном файле»)'):
            html = html.replace(t, '' if t.startswith(' ') else '')
    return Response(html, mimetype='text/html', headers={'Cache-Control': 'no-store'})
@app.route('/api/me')
def api_me(): return jsonify(user=session.get('u'), role=role())
@app.after_request
def nocache(r):
    r.headers.setdefault('Cache-Control', 'no-store')
    # сжатие ответов (данные пульта ~1 МБ): быстрее на телефоне
    if (not r.direct_passthrough and r.status_code == 200 and 'gzip' in request.headers.get('Accept-Encoding', '')
            and (r.mimetype in ('application/json', 'text/html', 'text/css', 'application/javascript', 'text/javascript')) and 'Content-Encoding' not in r.headers):
        data = r.get_data()
        if len(data) > 1500:
            import gzip as _gz
            r.set_data(_gz.compress(data, 5)); r.headers['Content-Encoding'] = 'gzip'; r.headers['Vary'] = 'Accept-Encoding'; r.headers['Content-Length'] = str(len(r.get_data()))
    return r
@app.route('/api/data')
def api_data():
    with db() as c: rows = c.execute('select * from cutdata order by cut').fetchall()
    M, S = [], []
    for r in rows: M += json.loads(r['metrics']); S += json.loads(r['series'])
    if role() == 'bank': M, S = bank_rows(M, S)
    return jsonify(metrics=uchet_walk(M), series=uchet_walk(S), status=None if role() == 'bank' else last_run(), measured={r['cut']: r['measured_at'] for r in rows})
@app.route('/api/status')
def api_status(): return jsonify(status=None if role() == 'bank' else last_run())
@app.route('/api/file')
def api_file():
    with db() as c: f = c.execute('select name, body from file where cut=? and doc=?', (request.args.get('cut', ''), request.args.get('doc', ''))).fetchone()
    if not f: return 'Файл для этого среза не собран', 404
    body = bytes(f['body'])
    if role() == 'bank':
        if request.args.get('doc', '') in BANK_FILES_BLOCK: return 'Нет доступа к этому файлу', 403
        body = bank_xlsx(request.args.get('doc', ''), body)
    body = uchet_xlsx(body)
    return Response(body, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': 'attachment; filename="linko_%s_%s.xlsx"; filename*=UTF-8\'\'%s' % (request.args.get('doc', ''), request.args.get('cut', ''), quote(f['name']))})
@app.route('/api/refresh', methods=['POST'])
def api_refresh():
    if role() in RO: return jsonify(error='У вас доступ только на просмотр'), 403
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
    force = request.args.get('force') == '1' and role() not in RO      # банк видит «Сейчас в учётной системе», но без пересчёта
    with TLOCK:
        if force or not TODAY['data'] or time.time() - TODAY['at'] > 600:
            try:
                r = subprocess.run([sys.executable, os.path.join('pult', 'today.py')], cwd=CALC, capture_output=True, text=True, timeout=90, env=dict(os.environ, PYTHONIOENCODING='utf-8'))
                if r.returncode != 0: raise RuntimeError((r.stderr or r.stdout)[-300:])
                TODAY.update(at=time.time(), data=json.loads(r.stdout), err=None)
            except Exception as e:
                TODAY['err'] = 'Не удалось получить данные из учётной системы: ' + re.sub(r'password=\S+', 'password=***', str(e))[-200:]
        d = dict(TODAY['data'] or {})
    with db() as c:
        row = c.execute('select cut, metrics from cutdata order by cut desc limit 1').fetchone()
    if row:
        last = {m['key']: m['value'] for m in json.loads(row['metrics'])}
        d['last_cut'] = row['cut']; d['last_subs'] = last.get('subs'); d['last_clients'] = last.get('clients')
    d['fetched_at'] = datetime.datetime.fromtimestamp(TODAY['at'], datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ') if TODAY['at'] else None
    d['error'] = TODAY['err']
    return jsonify(uchet_walk(d))
PLAN = {'busy': False}
PLOCK = threading.Lock()
def plan_dir():
    d = os.path.join(DATA, 'plan'); os.makedirs(d, exist_ok=True); return d
@app.route('/api/plan')
def api_plan():
    """План поступлений на 6 месяцев — считается на сегодня (pult/plan.py), хранится в data/plan/. Пересчёт: раз в 6 часов или по кнопке (не для просмотра)."""
    import time
    if role() == 'bank': return jsonify(error='Нет доступа'), 403
    fp = os.path.join(plan_dir(), 'plan.json'); force = request.args.get('force') == '1' and role() not in RO
    old = os.path.exists(fp) and (time.time() - os.path.getmtime(fp) > 6 * 3600 or json.load(open(fp, encoding='utf-8')).get('today') != datetime.date.today().isoformat())
    err = None
    if force or not os.path.exists(fp) or old:
        with PLOCK:
            fresh = os.path.exists(fp) and time.time() - os.path.getmtime(fp) < 60      # пока ждали — посчитал соседний запрос
            try:
                if fresh: raise StopIteration
                r = subprocess.run([sys.executable, os.path.join('pult', 'plan.py'), plan_dir()], cwd=CALC, capture_output=True, text=True, timeout=170, env=dict(os.environ, PYTHONIOENCODING='utf-8'))
                if r.returncode != 0: raise RuntimeError((r.stderr or r.stdout)[-300:])
            except StopIteration: pass
            except Exception as e:
                err = 'Не удалось пересчитать план из учётной системы: ' + re.sub(r'password=\S+', 'password=***', str(e))[-200:]
    if not os.path.exists(fp): return jsonify(error=err or 'План ещё не посчитан'), 500
    d = json.load(open(fp, encoding='utf-8')); d['computed_at'] = datetime.datetime.fromtimestamp(os.path.getmtime(fp), datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'); d['error'] = err
    return jsonify(uchet_walk(d))
@app.route('/api/plan/file')
def api_plan_file():
    if role() == 'bank': return 'Нет доступа', 403
    fp = os.path.join(plan_dir(), 'plan.json')
    if not os.path.exists(fp): return 'План ещё не посчитан', 404
    fn = json.load(open(fp, encoding='utf-8'))['file']; path = os.path.join(plan_dir(), fn)
    if not os.path.exists(path): return 'Файл не найден', 404
    return Response(uchet_xlsx(open(path, 'rb').read()), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': "attachment; filename=\"linko_plan.xlsx\"; filename*=UTF-8''%s" % quote(fn)})
@app.route('/api/history')
def api_history():
    if role() == 'bank': return jsonify(runs=[])
    with db() as c: rows = c.execute('select id, started, finished, state, cut, docs, msg, who from run order by id desc limit 50').fetchall()
    return jsonify(runs=[dict(r) for r in rows])
init()
import debts
debts.install(app, db, now, CFG, CALC)
