# -*- coding: utf-8 -*-
"""Обзвоны клиентов — сырая выгрузка из Odoo, в формате исходного файла (3 листа, формулы COUNTIFS)."""
import json, copy, datetime as dt, openpyxl
from cfg_cut import *
TPL = '_шаблоны_форматов/tpl_calls.xlsx'
rows = json.load(open(J('calls')))
N = len(rows); FIRST = 4; LAST = FIRST + N - 1
wb = openpyxl.load_workbook(TPL)
# ---------- лист 1: сырые данные ----------
ws = wb['Обзвоны (сырые данные)']
old_last = 1370
proto = [copy.copy(ws.cell(4, c)._style) for c in range(1, 29)]
date_style = copy.copy(ws.cell(4, 2)._style)
note_style = copy.copy(ws.cell(1372, 1)._style)
for r in range(4, ws.max_row + 1):
    for c in range(1, 29):
        ws.cell(r, c).value = None
def d(v): return dt.datetime.strptime(v, '%Y-%m-%d') if v else None
DATE_COLS = (2, 23, 24, 25)
for k, row in enumerate(rows):
    r = FIRST + k
    for c in range(1, 29):
        v = row[c - 1]
        cell = ws.cell(r, c)
        if c in DATE_COLS:
            cell.value = d(v); cell._style = copy.copy(date_style)
        else:
            cell.value = v; cell._style = copy.copy(proto[c - 1])
    ws.row_dimensions[r].height = 15
hist = sum(1 for x in rows if x[23])
notes = [
 'Источник: Odoo linktrade (odoo.linko.uz), модель x_nps «NPS/обзвон клиентов». Выгрузка на %s, звонки с датой по %s включительно. Записи, помеченные в Odoo как архивные, включены (колонка «Активная запись»).' % (EXPORT_RU, CUT_RU),
 'Телефоны контактных лиц обезличены (последние 4 цифры). Названия клиентов оставлены как есть. Комментарии операторов приведены дословно, без правки языка и орфографии.',
 '«Дата закрытия (из истории)» заполнена только там, где смена статуса попала в журнал изменений Odoo (%d записи, сен–окт 2025). Для остальных закрытых обращений ориентир — «Дата последнего изменения».' % hist]
for i, t in enumerate(notes):
    c = ws.cell(LAST + 2 + i, 1); c.value = t; c._style = copy.copy(note_style)
ws.auto_filter.ref = 'A3:AB%d' % LAST
# ---------- лист 2: сводка ----------
ws2 = wb['Сводка']
S = lambda r: [copy.copy(ws2.cell(r, c)._style) for c in range(1, 8)]
st_m, st_nd, st_tot = S(4), S(16), S(17)
st_sec, st_h, st_r, st_t = S(20), S(21), S(22), S(27)
st_sec2, st_h2, st_r2, st_t2 = S(30), S(31), S(32), S(38)
dest = [ws2.cell(r, 1).value for r in range(22, 27)]
stat = [ws2.cell(r, 1).value for r in range(32, 37)]
sec1, sec2 = ws2.cell(20, 1).value, ws2.cell(30, 1).value
for r in range(4, ws2.max_row + 1):
    for c in range(1, 8):
        ws2.cell(r, c).value = None; ws2.cell(r, c)._style = copy.copy(ws2.cell(2, c)._style)
months = sorted({x[2] for x in rows if x[2]})
y, m = 2025, 9; allm = []
while '%04d-%02d' % (y, m) <= months[-1]:
    allm.append('%04d-%02d' % (y, m)); m += 1
    if m == 13: y, m = y + 1, 1
RAW = "'Обзвоны (сырые данные)'!"
def put(r, vals, styles):
    for c, v in enumerate(vals, 1):
        ws2.cell(r, c).value = v
    for c in range(1, 8):
        ws2.cell(r, c)._style = copy.copy(styles[c - 1])
    ws2.row_dimensions[r].height = 15
r = 4
for mm in allm:
    put(r, [mm] + ["=COUNTIFS(%s$C$%d:$C$%d,$A%d,%s$I$%d:$I$%d,%s$3)" % (RAW, FIRST, LAST, r, RAW, FIRST, LAST, col) for col in 'BCDE']
        + ['=SUM(B%d:E%d)' % (r, r), '=IF(F%d=0,0,C%d/F%d)' % (r, r, r)], st_m); r += 1
m_first, m_last = 4, r - 1
put(r, ['(без даты)'] + ["=COUNTIF(%s$I$%d:$I$%d,%s$3)-SUM(%s%d:%s%d)" % (RAW, FIRST, LAST, col, col, m_first, col, m_last) for col in 'BCDE']
    + ['=SUM(B%d:E%d)' % (r, r), '=IF(F%d=0,0,C%d/F%d)' % (r, r, r)], st_nd); nd = r; r += 1
put(r, ['ИТОГО'] + ['=SUM(%s%d:%s%d)' % (col, m_first, col, nd) for col in 'BCDEF'] + ['=IF(F%d=0,0,C%d/F%d)' % (r, r, r)], st_tot); tot_row = r
r += 3
put(r, [sec1], st_sec); r += 1
put(r, ['Куда направлена', 'Кол-во', 'Доля'], st_h); r += 1
t = r + len(dest)
for name in dest:
    put(r, [name, "=COUNTIFS(%s$L$%d:$L$%d,$A%d)" % (RAW, FIRST, LAST, r), '=IF($B$%d=0,0,B%d/$B$%d)' % (t, r, t)], st_r); r += 1
put(r, ['ИТОГО', '=SUM(B%d:B%d)' % (t - len(dest), t - 1)], st_t)
r += 3
put(r, [sec2], st_sec2); r += 1
put(r, ['Статус', 'Кол-во', 'Доля'], st_h2); r += 1
t = r + len(stat) + 1; s0 = r
for name in stat:
    put(r, [name, "=COUNTIFS(%s$O$%d:$O$%d,$A%d)" % (RAW, FIRST, LAST, r), '=IF($B$%d=0,0,B%d/$B$%d)' % (t, r, t)], st_r2); r += 1
put(r, ['Не заполнен', '=%d-SUM(B%d:B%d)' % (N, s0, r - 1), '=IF($B$%d=0,0,B%d/$B$%d)' % (t, r, t)], st_r2); r += 1
put(r, ['ИТОГО', '=SUM(B%d:B%d)' % (s0, r - 1)], st_t2)
# ---------- лист 3: методика ----------
ws3 = wb['Методика и справочник']
dd = [x[1] for x in rows if x[1]]
f = lambda s: '%s.%s.%s' % (s[8:10], s[5:7], s[0:4])
ws3['B3'] = 'Odoo linktrade (odoo.linko.uz), модель x_nps — «NPS / обзвон клиентов». Выгружены все записи модели с датой звонка по %s включительно (и записи без даты), включая архивные.' % CUT_RU
ws3['B4'] = EXPORT_RU
ws3['B5'] = str(N)
ws3['B6'] = '%s — %s' % (f(min(dd)), f(max(dd)))
ws3['B7'] = 'Данных за 2024 год в Odoo нет: модель обзвонов заведена в сентябре 2025 г., первая запись — %s. Всё, что было до этого, в системе не фиксировалось.' % f(min(dd))
ws3['B19'] = 'Сотрудник, назначенный по обращению, и его подразделение по кадровой карточке (для уволенных — по последней, архивной карточке). Отдел пуст, если у пользователя нет карточки сотрудника.'
ws3['B24'] = 'Восстановлена из журнала изменений Odoo по факту перевода в «Решено»/«Отмена». Такие записи в журнале есть только за сен–окт 2025 (%d шт.) — позже этот статус перестали трекать.' % hist
fn = OUTP('Обзвоны_клиентов_сырая_выгрузка_Odoo_%s.xlsx' % SUF)
wb.save(fn)
import collections
def extra():
    """Охват базы, открытые проблемы, операторы, оценки — для пульта."""
    pid = json.load(open(J('calls_pid'))); base = json.load(open(J('base'), encoding='utf-8'))
    bp = {b['partner_id'] for b in base}; c90 = (dt.date.fromisoformat(CUT) - dt.timedelta(days=90)).isoformat()
    reached = {pid.get(str(x[0])) for x in rows if x[1] and c90 < x[1] <= CUT and x[8] in ('Всё ок', 'Проблема')} & bp
    tried = {pid.get(str(x[0])) for x in rows if x[1] and c90 < x[1] <= CUT} & bp
    openp = [x for x in rows if x[8] == 'Проблема' and x[14] not in ('Решено', 'Отмена')]
    ops = collections.Counter(x[21] or 'не указан' for x in rows if x[2] == LAST_M)
    sc = [x[17] for x in rows if x[17] and x[1] and x[1] <= CUT]
    return {'base_clients': len(bp), 'reached90': len(reached), 'tried90': len(tried), 'open_problems': len(openp),
            'open_by_route': collections.Counter(x[11] or 'не направлена' for x in openp), 'ops': ops, 'scored': len(sc), 'scored_month': sum(1 for x in rows if x[17] and x[2] == LAST_M)}
res = collections.Counter((x[2], x[8]) for x in rows)
bym = {mm: {k: res.get((mm, k), 0) for k in ('Всё ок', 'Проблема', 'Не дозвонились', 'Нет данных')} for mm in allm}
json.dump({'N': N, 'bym': bym, 'route': collections.Counter(x[11] for x in rows if x[11]), 'stage': collections.Counter(x[14] or '-' for x in rows),
           'prod': collections.Counter(x[5] for x in rows), **extra()}, open(J('res_calls'), 'w'), ensure_ascii=False)
print(fn, N, LAST, 'месяцев', len(allm))
for mm in allm[-4:]: print(mm, bym[mm])
