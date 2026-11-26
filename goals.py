# -*- coding: utf-8 -*-
"""Раздел «Цели и планы» → «Планы продаж» (26.11.2026).
План: сумма нового MRR в сумах на месяц и продукт (SFA, POS, AI ассистент, Perfect store); горизонт — до 24 месяцев вперёд.
Факт: calc/pult/sales_fact.py — новый MRR месяца + расширения/сокращения подписок младше 6 месяцев (журнал MRR учётной системы).
Права: ввод плана — только admin; просмотр — admin и viewer; роль bank — нет доступа (guard в app.py)."""
import os, sys, json, time, datetime, subprocess, threading, re
from flask import Blueprint, request, session, jsonify
bp = Blueprint('goals', __name__)
G = {}; FACT = {'at': 0, 'data': None, 'err': None}; FLOCK = threading.Lock()
PRODS = ['SFA', 'POS', 'AI', 'PS']; PNAME = {'SFA': 'SFA', 'POS': 'POS', 'AI': 'AI ассистент', 'PS': 'Perfect store'}
FIRST = '2026-01'; HORIZON = 24
def install(app, db, now, cfg, calc, role):
    G.update(db=db, now=now, cfg=cfg, calc=calc, role=role)
    with db() as c:
        c.executescript("""
        create table if not exists sales_plan(ym text not null, prod text not null, amount real not null, updated_at text, updated_by text, primary key(ym, prod));
        create table if not exists sales_plan_log(id integer primary key autoincrement, at text, who text, ym text, prod text, old real, new real);
        """)
    app.register_blueprint(bp)
def ym_add(ym, n):
    y, m = int(ym[:4]), int(ym[5:7]) - 1 + n
    return '%04d-%02d' % (y + m // 12, m % 12 + 1)
def bounds():
    cur = datetime.date.today().strftime('%Y-%m'); return FIRST, ym_add(cur, HORIZON - 1), cur
def fact(force=False):
    with FLOCK:
        if force or not FACT['data'] or time.time() - FACT['at'] > 600:
            try:
                r = subprocess.run([sys.executable, os.path.join('pult', 'sales_fact.py')], cwd=G['calc'], capture_output=True, text=True, timeout=120, env=dict(os.environ, PYTHONIOENCODING='utf-8'))
                if r.returncode != 0: raise RuntimeError((r.stderr or r.stdout)[-300:])
                FACT.update(at=time.time(), data=json.loads(r.stdout), err=None)
            except Exception as e:
                FACT['err'] = 'Не удалось получить факт продаж из учётной системы: ' + re.sub(r'password=\S+', 'password=***', str(e))[-200:]
        return FACT['data'], FACT['err']
@bp.route('/api/goals/sales')
def sales_get():
    force = request.args.get('force') == '1' and G['role']() == 'admin'
    data, err = fact(force)
    with G['db']() as c:
        plans = [dict(r) for r in c.execute('select ym, prod, amount, updated_at, updated_by from sales_plan order by ym, prod')]
        log = [dict(r) for r in c.execute('select at, who, ym, prod, old, new from sales_plan_log order by id desc limit 100')]
    lo, hi, cur = bounds()
    return jsonify(plans=plans, log=log, fact=(data or {}).get('fact', {}), details=(data or {}).get('details', {}), computed=(data or {}).get('computed'),
                   error=err, can_edit=G['role']() == 'admin', first=lo, last=hi, current=cur, prods=PRODS, pname=PNAME, window=(data or {}).get('window', 6))
@bp.route('/api/goals/sales', methods=['POST'])
def sales_set():
    if G['role']() != 'admin': return jsonify(error='Планы может менять только администратор'), 403
    j = request.get_json(silent=True) or {}
    ym, prod, amt = str(j.get('ym', '')), str(j.get('prod', '')), j.get('amount')
    lo, hi, cur = bounds()
    if not re.fullmatch(r'20\d\d-(0[1-9]|1[0-2])', ym): return jsonify(error='Период — год и месяц'), 400
    if ym < lo or ym > hi: return jsonify(error='Период — с %s по %s (горизонт планирования 2 года)' % (lo, hi)), 400
    if prod not in PRODS: return jsonify(error='Неизвестный продукт'), 400
    if amt in (None, ''): amt = None
    else:
        try: amt = round(float(amt), 2)
        except (TypeError, ValueError): return jsonify(error='Сумма — число в сумах'), 400
        if amt < 0 or amt > 1e13: return jsonify(error='Сумма должна быть от 0'), 400
    with G['db']() as c:
        old = c.execute('select amount from sales_plan where ym=? and prod=?', (ym, prod)).fetchone()
        old = old[0] if old else None
        if amt is None: c.execute('delete from sales_plan where ym=? and prod=?', (ym, prod))
        else: c.execute('insert or replace into sales_plan(ym, prod, amount, updated_at, updated_by) values(?,?,?,?,?)', (ym, prod, amt, G['now'](), session.get('u')))
        if old != amt: c.execute('insert into sales_plan_log(at, who, ym, prod, old, new) values(?,?,?,?,?,?)', (G['now'](), session.get('u'), ym, prod, old, amt))
    return jsonify(ok=True)
