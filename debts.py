# -*- coding: utf-8 -*-
"""Учёт долгов компании — ведётся вручную на пульте, Odoo не используется.
Долг = карточка (название, у кого взяли, тип, валюта, дата, срок) + операции: «взяли / добавили к долгу» и «выплата».
Операции и долги не удаляются — только аннулируются с причиной. Каждое изменение пишется в журнал (кто и когда).
Доступ: data/config.json → "perms": {"логин": {"debts": "edit" | "view" | "none"}}; если не указано — "edit"."""
import os, sys, io, json, datetime
from urllib.parse import quote
from flask import Blueprint, request, session, jsonify, Response
RATE = 12000
KINDS = ['Займ', 'Поставщик / услуги', 'Аренда', 'Зарплата / сотрудники', 'Учредитель', 'Налоги и сборы', 'Прочее']
CTYPES = ['Физлицо', 'Организация']
CUR = {'UZS': 'сум', 'USD': '$'}
bp = Blueprint('debts', __name__)
G = {}
def install(app, db, now, cfg, calc):
    G.update(db=db, now=now, cfg=cfg, calc=calc)
    with db() as c:
        c.executescript("""
        create table if not exists debt(id integer primary key autoincrement, title text not null, creditor text not null, ctype text, kind text,
          currency text not null, opened text not null, due text, note text, void integer default 0, void_reason text, void_at text, void_by text,
          created_at text, created_by text, updated_at text, updated_by text);
        create table if not exists debt_op(id integer primary key autoincrement, debt_id integer not null references debt(id), op text not null check(op in ('take','pay')),
          date text not null, amount real not null, note text, void integer default 0, void_reason text, void_at text, void_by text, created_at text, created_by text);
        create index if not exists debt_op_debt on debt_op(debt_id);
        create table if not exists debt_log(id integer primary key autoincrement, at text, who text, debt_id integer, op_id integer, action text, detail text);
        """)
    app.register_blueprint(bp)
def perm():
    return ((G['cfg'].get('perms') or {}).get(session.get('u') or '', {}) or {}).get('debts', 'edit')
class Bad(Exception): pass
@bp.errorhandler(Bad)
def bad(e): return jsonify(error=str(e)), 400
def need(level):
    p = perm()
    if p == 'none' or (level == 'edit' and p != 'edit'): raise PermissionError
@bp.errorhandler(PermissionError)
def denied(e): return jsonify(error='Нет доступа к этому действию'), 403
today = lambda: datetime.date.today().isoformat()
def txt(j, k, req=False, n=200):
    v = str(j.get(k) or '').strip()
    if req and not v: raise Bad({'title': 'Укажите название долга', 'creditor': 'Укажите, у кого взяли (ФИО или организация)', 'reason': 'Укажите причину'}.get(k, 'Заполните поле ' + k))
    if len(v) > n: raise Bad('Слишком длинный текст (%s): не больше %d символов' % (k, n))
    return v
def dt(j, k, req=True, future=False):
    v = str(j.get(k) or '').strip()
    if not v:
        if req: raise Bad('Укажите дату')
        return None
    try: d = datetime.date.fromisoformat(v)
    except ValueError: raise Bad('Дата в неверном формате: %s' % v)
    if d.year < 2015 or d.year > 2100: raise Bad('Проверьте год в дате: %s' % v)
    if not future and v > today(): raise Bad('Дата операции не может быть в будущем')
    return v
def amt(j, cur):
    raw = str(j.get('amount') or '').replace(' ', '').replace(' ', '').replace(',', '.')
    try: a = round(float(raw), 2 if cur == 'USD' else 0)
    except ValueError: raise Bad('Сумма должна быть числом')
    if not (a > 0): raise Bad('Сумма должна быть больше нуля')
    if a > 1e13: raise Bad('Слишком большая сумма')
    return a
def money(a, cur):
    s = '{:,.2f}'.format(a).rstrip('0').rstrip('.') if cur == 'USD' else '{:,.0f}'.format(a)
    s = s.replace(',', ' ')
    return '$' + s if cur == 'USD' else s + ' сум'
ru = lambda d: d[8:10] + '.' + d[5:7] + '.' + d[:4] if d else '—'
def log(c, debt_id, op_id, action, detail):
    c.execute('insert into debt_log(at, who, debt_id, op_id, action, detail) values(?,?,?,?,?,?)', (G['now'](), session.get('u'), debt_id, op_id, action, detail))
def load(c):
    D = [dict(r) for r in c.execute('select * from debt order by opened desc, id desc')]
    ops = {}
    for r in c.execute('select * from debt_op order by date, id'): ops.setdefault(r['debt_id'], []).append(dict(r))
    t = today()
    for d in D:
        d['ops'] = ops.get(d['id'], [])
        live = [o for o in d['ops'] if not o['void']]
        d['taken'] = sum(o['amount'] for o in live if o['op'] == 'take'); d['paid'] = sum(o['amount'] for o in live if o['op'] == 'pay')
        d['left'] = round(d['taken'] - d['paid'], 2); k = RATE if d['currency'] == 'USD' else 1
        d['taken_uzs'], d['paid_uzs'], d['left_uzs'] = d['taken'] * k, d['paid'] * k, d['left'] * k
        d['last_pay'] = max((o['date'] for o in live if o['op'] == 'pay'), default=None)
        d['status'] = 'void' if d['void'] else 'closed' if d['left'] <= 0.004 else 'overdue' if d['due'] and d['due'] < t else 'open'
    return D
def one(c, i):
    r = c.execute('select * from debt where id=?', (i,)).fetchone()
    if not r: raise Bad('Долг не найден')
    return r
def left_of(c, i):
    r = c.execute("select coalesce(sum(case when op='take' then amount else -amount end),0) from debt_op where debt_id=? and void=0", (i,)).fetchone()
    return round(r[0], 2)
@bp.route('/api/debts')
def api_list():
    need('view')
    with G['db']() as c:
        D = load(c); L = [dict(r) for r in c.execute('select * from debt_log order by id desc limit 300')]
    return jsonify(can=perm(), rate=RATE, kinds=KINDS, ctypes=CTYPES, today=today(), debts=D, log=L)
@bp.route('/api/debts', methods=['POST'])
def api_create():
    need('edit'); j = request.get_json(silent=True) or {}
    cur = j.get('currency') if j.get('currency') in CUR else 'UZS'
    title, cred = txt(j, 'title', True), txt(j, 'creditor', True)
    ctype = j.get('ctype') if j.get('ctype') in CTYPES else None; kind = j.get('kind') if j.get('kind') in KINDS else 'Прочее'
    opened, due, a, note = dt(j, 'date'), dt(j, 'due', False, True), amt(j, cur), txt(j, 'note', n=1000)
    if due and due < opened: raise Bad('Срок возврата раньше даты долга')
    n, u = G['now'](), session.get('u')
    with G['db']() as c:
        i = c.execute('insert into debt(title, creditor, ctype, kind, currency, opened, due, note, created_at, created_by, updated_at, updated_by) values(?,?,?,?,?,?,?,?,?,?,?,?)',
                      (title, cred, ctype, kind, cur, opened, due, note, n, u, n, u)).lastrowid
        o = c.execute("insert into debt_op(debt_id, op, date, amount, note, created_at, created_by) values(?,?,?,?,?,?,?)", (i, 'take', opened, a, 'Сумма долга при заведении', n, u)).lastrowid
        log(c, i, o, 'create', 'Новый долг «%s» — %s, %s от %s' % (title, cred, money(a, cur), ru(opened)))
    return jsonify(ok=True, id=i)
FIELDS = [('title', 'Название'), ('creditor', 'У кого взяли'), ('ctype', 'Физлицо / организация'), ('kind', 'Тип'), ('due', 'Срок возврата'), ('note', 'Комментарий'), ('opened', 'Дата долга')]
@bp.route('/api/debts/<int:i>', methods=['POST'])
def api_edit(i):
    need('edit'); j = request.get_json(silent=True) or {}
    with G['db']() as c:
        d = dict(one(c, i))
        if d['void']: raise Bad('Долг аннулирован — изменить нельзя')
        new = {'title': txt(j, 'title', True), 'creditor': txt(j, 'creditor', True), 'ctype': j.get('ctype') if j.get('ctype') in CTYPES else None,
               'kind': j.get('kind') if j.get('kind') in KINDS else 'Прочее', 'due': dt(j, 'due', False, True), 'note': txt(j, 'note', n=1000), 'opened': dt(j, 'date')}
        if new['due'] and new['due'] < new['opened']: raise Bad('Срок возврата раньше даты долга')
        ch = ['%s: «%s» → «%s»' % (lab, (ru(d[k]) if k in ('due', 'opened') else d[k]) or '—', (ru(new[k]) if k in ('due', 'opened') else new[k]) or '—') for k, lab in FIELDS if (d[k] or None) != (new[k] or None)]
        if not ch: return jsonify(ok=True, same=True)
        first = c.execute("select id from debt_op where debt_id=? and op='take' order by id limit 1", (i,)).fetchone()
        early = c.execute('select min(date) from debt_op where debt_id=? and void=0 and id<>?', (i, first[0] if first else 0)).fetchone()[0]
        if early and early < new['opened']: raise Bad('Есть операции раньше новой даты долга (%s) — сначала проверьте их' % ru(early))
        if first and new['opened'] != d['opened']: c.execute('update debt_op set date=? where id=?', (new['opened'], first[0]))
        c.execute('update debt set title=?, creditor=?, ctype=?, kind=?, due=?, note=?, opened=?, updated_at=?, updated_by=? where id=?',
                  (new['title'], new['creditor'], new['ctype'], new['kind'], new['due'], new['note'], new['opened'], G['now'](), session.get('u'), i))
        log(c, i, None, 'edit', 'Изменён долг «%s»: %s' % (new['title'], '; '.join(ch)))
    return jsonify(ok=True)
@bp.route('/api/debts/<int:i>/void', methods=['POST'])
def api_void(i):
    need('edit'); j = request.get_json(silent=True) or {}; reason = txt(j, 'reason', True, 500)
    with G['db']() as c:
        d = one(c, i)
        if d['void']: raise Bad('Долг уже аннулирован')
        c.execute('update debt set void=1, void_reason=?, void_at=?, void_by=? where id=?', (reason, G['now'](), session.get('u'), i))
        log(c, i, None, 'void', 'Аннулирован долг «%s» — %s. Причина: %s' % (d['title'], d['creditor'], reason))
    return jsonify(ok=True)
@bp.route('/api/debts/<int:i>/ops', methods=['POST'])
def api_op(i):
    need('edit'); j = request.get_json(silent=True) or {}
    op = j.get('op')
    if op not in ('take', 'pay'): raise Bad('Неизвестная операция')
    with G['db']() as c:
        d = one(c, i)
        if d['void']: raise Bad('Долг аннулирован — операции по нему вносить нельзя')
        date, a, note = dt(j, 'date'), amt(j, d['currency']), txt(j, 'note', n=500)
        if date < d['opened']: raise Bad('Дата операции раньше даты долга (%s)' % ru(d['opened']))
        if op == 'pay':
            lf = left_of(c, i)
            if a > lf + 0.004: raise Bad('Выплата больше остатка долга: осталось %s' % money(lf, d['currency']))
        o = c.execute('insert into debt_op(debt_id, op, date, amount, note, created_at, created_by) values(?,?,?,?,?,?,?)', (i, op, date, a, note, G['now'](), session.get('u'))).lastrowid
        log(c, i, o, op, '%s по долгу «%s»: %s от %s%s' % ('Выплата' if op == 'pay' else 'Добавлено к долгу', d['title'], money(a, d['currency']), ru(date), ('. ' + note) if note else ''))
    return jsonify(ok=True, id=o)
@bp.route('/api/debt_ops/<int:o>/void', methods=['POST'])
def api_op_void(o):
    need('edit'); j = request.get_json(silent=True) or {}; reason = txt(j, 'reason', True, 500)
    with G['db']() as c:
        r = c.execute('select * from debt_op where id=?', (o,)).fetchone()
        if not r: raise Bad('Операция не найдена')
        if r['void']: raise Bad('Операция уже аннулирована')
        d = one(c, r['debt_id'])
        if r['op'] == 'take' and left_of(c, d['id']) - r['amount'] < -0.004:
            raise Bad('После этого выплаты окажутся больше долга. Сначала аннулируйте лишние выплаты.')
        c.execute('update debt_op set void=1, void_reason=?, void_at=?, void_by=? where id=?', (reason, G['now'](), session.get('u'), o))
        log(c, d['id'], o, 'op_void', 'Аннулирована операция по долгу «%s»: %s %s от %s. Причина: %s' % (d['title'], 'выплата' if r['op'] == 'pay' else 'начисление', money(r['amount'], d['currency']), ru(r['date']), reason))
    return jsonify(ok=True)
@bp.route('/api/debts.xlsx')
def api_xlsx():
    need('view')
    if G['calc'] not in sys.path: sys.path.insert(0, G['calc'])
    import openpyxl
    from openpyxl.utils import get_column_letter as L
    from style_lib import title, header, datarow, totalrow, notes, widths, NUM
    with G['db']() as c:
        D = load(c); LG = [dict(r) for r in c.execute('select * from debt_log order by id')]
    st = {'open': 'Открыт', 'overdue': 'Просрочен', 'closed': 'Погашен', 'void': 'Аннулирован'}
    D.sort(key=lambda d: ({'overdue': 0, 'open': 1, 'closed': 2, 'void': 3}[d['status']], d['opened']))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Долги'; t = today()
    title(ws, 'A1', 'Долги компании на %s' % ru(t))
    header(ws, 3, ['№', 'Название', 'У кого взяли', 'Физлицо / организация', 'Тип', 'Валюта', 'Дата долга', 'Срок возврата', 'Было (взято)', 'Вернули', 'Осталось',
                   'Осталось, сум (по 12 000)', 'Последняя выплата', 'Статус', 'Комментарий'])
    r = 4; DF = 'DD.MM.YYYY'; D_ = lambda s: datetime.date.fromisoformat(s) if s else None
    for n, d in enumerate(D, 1):
        if d['void']: continue
        cf = '"$"#,##0.00' if d['currency'] == 'USD' else NUM
        datarow(ws, r, [n, d['title'], d['creditor'], d['ctype'] or '', d['kind'] or '', d['currency'], D_(d['opened']), D_(d['due']), d['taken'], d['paid'], '=I%d-J%d' % (r, r),
                        '=K%d*%d' % (r, RATE if d['currency'] == 'USD' else 1), D_(d['last_pay']), st[d['status']], d['note'] or ''],
                fmts=[None, None, None, None, None, None, DF, DF, cf, cf, cf, NUM, DF, None, None], blue_inputs=False); r += 1
    totalrow(ws, r, ['ИТОГО должны, сум', None, None, None, None, None, None, None, None, None, None, '=SUM(L4:L%d)' % (r - 1)], fmts=[None] * 11 + [NUM])
    notes(ws, r + 2, ['Источник: учёт долгов на пульте Linko (ведётся вручную, Odoo не используется). Выгрузка %s.' % ru(t),
                      'Долги в долларах в колонке «Осталось, сум» пересчитаны по 12 000 сум/$. Аннулированные долги не показаны (они есть в журнале).',
                      '«Было» — сумма долга при заведении плюс всё, что добавляли к долгу позже; «Вернули» — все выплаты.'])
    widths(ws, {'A': 5, 'B': 30, 'C': 28, 'D': 14, 'E': 18, 'F': 8, 'G': 12, 'H': 12, 'I': 16, 'J': 16, 'K': 16, 'L': 18, 'M': 12, 'N': 12, 'O': 40}); ws.freeze_panes = 'C4'
    ws = wb.create_sheet('Операции'); title(ws, 'A1', 'Все операции по долгам')
    header(ws, 3, ['Дата', 'Долг', 'У кого взяли', 'Операция', 'Сумма', 'Валюта', 'Сумма, сум (по 12 000)', 'Комментарий', 'Внёс', 'Когда внесено', 'Аннулирована', 'Причина аннулирования'])
    r = 4; ops = sorted(((o, d) for d in D for o in d['ops']), key=lambda x: (x[0]['date'], x[0]['id']))
    for o, d in ops:
        sg = 1 if o['op'] == 'take' else -1; k = RATE if d['currency'] == 'USD' else 1
        datarow(ws, r, [D_(o['date']), d['title'], d['creditor'], 'Взяли / добавили к долгу' if o['op'] == 'take' else 'Выплата', sg * o['amount'], d['currency'], sg * o['amount'] * k,
                        o['note'] or '', o['created_by'] or '', (o['created_at'] or '').replace('T', ' ').replace('Z', ' UTC'), 'да' if (o['void'] or d['void']) else '', o['void_reason'] or (d['void_reason'] if d['void'] else '') or ''],
                fmts=[DF, None, None, None, '#,##0.##', None, NUM], blue_inputs=False); r += 1
    widths(ws, {'A': 12, 'B': 30, 'C': 28, 'D': 24, 'E': 16, 'F': 8, 'G': 18, 'H': 36, 'I': 10, 'J': 20, 'K': 12, 'L': 36}); ws.freeze_panes = 'A4'
    if r > 4: ws.auto_filter.ref = 'A3:L%d' % (r - 1)
    ws = wb.create_sheet('Журнал изменений'); title(ws, 'A1', 'Журнал изменений')
    header(ws, 3, ['Когда (UTC)', 'Кто', 'Что сделано'])
    for k, g in enumerate(LG, 4): datarow(ws, k, [(g['at'] or '').replace('T', ' ').replace('Z', ''), g['who'] or '', g['detail']], blue_inputs=False)
    widths(ws, {'A': 20, 'B': 12, 'C': 120})
    b = io.BytesIO(); wb.save(b); fn = 'Долги_на_%s.xlsx' % ru(t)
    return Response(b.getvalue(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': 'attachment; filename="linko_debts.xlsx"; filename*=UTF-8\'\'%s' % quote(fn)})
