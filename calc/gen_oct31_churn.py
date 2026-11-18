# -*- coding: utf-8 -*-
"""Отток подписок: причины и сегменты — в формате исходного файла (4 листа на формулах)."""
import json, copy, collections, openpyxl
from cfg_cut import *
TPL = '_шаблоны_форматов/tpl_churn.xlsx'
d = json.load(open(J('churn'))); c = json.load(open(J('const'))); mr = json.load(open(J('mrr')))
# причина Odoo -> (название в отчёте)
NAME = {3: 'Клиент не платит и не отвечает', 19: 'Недоступен', 15: 'Клиент отказался по собственным причинам', 8: 'Клиент отказался по собственным причинам',
        4: 'Клиент отказался по собственным причинам', 6: 'Клиент отказался по собственным причинам', 11: 'Клиент отказался по собственным причинам',
        27: 'Временная приостановка', 25: 'Закрытие бизнеса клиента', 21: 'Кадровые проблемы клиента', 26: 'Отсутствие товара', 24: 'Сезонная пауза',
        18: 'Отказ платить', 1: 'Отказ платить', 12: 'Не оплачивают', 9: 'Не оплачивают', 10: 'Не оплачивают', 17: 'Неплатёжеспособность',
        22: 'Сложность продукта', 2: 'Сложность продукта', 16: 'Клиент не пользовался системой', 5: 'Клиент не пользовался системой',
        23: 'Технические проблемы внедрения', 13: 'Неправильно сделали продажу', 28: 'Нецелевой клиент', 20: 'Ушёл к конкуренту',
        14: 'Клиент выбрал другую систему', 29: 'Клиент открыл новый сервер', 7: 'Клиент открыл новый сервер', None: 'Причина не указана'}
wb = openpyxl.load_workbook(TPL)
ws = wb['Причины по продуктам']
ORDER = [(ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value) for r in range(4, 23)]
ORDER.append(('Причина не указана', 'Причина не указана', 'Нет данных'))
agg = collections.defaultdict(lambda: [0, 0.0]); partners = {'POS': set(), 'SFA': set()}
for r in d['closed']:
    nm = NAME[r['cr']]
    agg[(nm, r['prod'])][0] += r['n']; agg[(nm, r['prod'])][1] += r['m']; partners[r['prod']].update(r['partners'])
assert all(k[0] in [o[0] for o in ORDER] for k in agg)
def styles(w, r): return [copy.copy(w.cell(r, k)._style) for k in range(1, w.max_column + 1)]
def put(w, r, vals, st):
    for k in range(1, len(st) + 1): w.cell(r, k)._style = copy.copy(st[k - 1])
    for k, v in enumerate(vals, 1): w.cell(r, k).value = v
# ---------- лист 2: причины ----------
s_r, s_t = styles(ws, 4), styles(ws, 23)
r = 4
for nm, cat, man in ORDER:
    put(ws, r, [nm, cat, man, agg[(nm, 'POS')][0], agg[(nm, 'POS')][1], agg[(nm, 'SFA')][0], agg[(nm, 'SFA')][1], '=D%d+F%d' % (r, r), '=E%d+G%d' % (r, r)], s_r); r += 1
put(ws, r, ['ИТОГО', None, None] + ['=SUM(%s4:%s%d)' % (k, k, r - 1) for k in 'DEFGHI'], s_t)
ws['A1'] = 'Причины оттока: POS и SFA раздельно (закрытия с 01.02.2026 по %s)' % CUT_RU
# ---------- лист 3: категории ----------
ws = wb['Категории причин']
CATS = [(ws.cell(r, 1).value, ws.cell(r, 2).value) for r in range(4, 11)] + [('Причина не указана', 'Нет данных')]
s_r, s_t = styles(ws, 4), styles(ws, 11)
cat_of = {o[0]: o[1] for o in ORDER}
ca = collections.defaultdict(lambda: [0, 0.0])
for (nm, p), (n, m) in agg.items(): ca[(cat_of[nm], p)][0] += n; ca[(cat_of[nm], p)][1] += m
r = 4; tr = 4 + len(CATS)
for cat, man in CATS:
    put(ws, r, [cat, man, ca[(cat, 'POS')][0], ca[(cat, 'POS')][1], ca[(cat, 'SFA')][0], ca[(cat, 'SFA')][1], '=C%d+E%d' % (r, r), '=D%d+F%d' % (r, r), '=H%d/$H$%d' % (r, tr)], s_r); r += 1
put(ws, r, ['ИТОГО', None] + ['=SUM(%s4:%s%d)' % (k, k, r - 1) for k in 'CDEFGH'] + ['=H%d/H%d' % (r, r)], s_t)
ws['A1'] = 'Отток по категориям причин (с 01.02.2026 по %s)' % CUT_RU
# ---------- лист 1: свод ----------
ws = wb['Свод по оттоку']
tot = {p: [sum(v[0] for k, v in agg.items() if k[1] == p), sum(v[1] for k, v in agg.items() if k[1] == p)] for p in ('POS', 'SFA')}
S = c['SPLIT']
ws['C4'], ws['D4'] = tot['POS'][0], tot['SFA'][0]
ws['C5'], ws['D5'] = tot['POS'][1], tot['SFA'][1]
ws['C8'], ws['D8'] = len(partners['POS']), len(partners['SFA'])
ws['C10'], ws['D10'] = S['POS']['3_progress'][0], S['SFA']['3_progress'][0]
ws['C11'], ws['D11'] = S['POS']['4_paused'][0], S['SFA']['4_paused'][0]
ws['C13'], ws['D13'] = S['POS']['3_progress'][1], S['SFA']['3_progress'][1]
ws['C14'], ws['D14'] = S['POS']['4_paused'][1], S['SFA']['4_paused'][1]
ws['C18'], ws['D18'] = agg[('Клиент открыл новый сервер', 'POS')][0], agg[('Клиент открыл новый сервер', 'SFA')][0]
top = sorted(((v[1], k[0], k[1], v[0]) for k, v in agg.items()), reverse=True)[:5]
for i, (m, nm, p, n) in enumerate(top):
    r = 23 + i
    ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value = nm, p, n, m
    ws.cell(r, 5).value = '=D%d/$B$5' % r
ws['A1'] = 'Linktrade — отток подписок: сегменты и причины (с февраля 2026, срез %s)' % CUT_RU
# ---------- лист 4: динамика ----------
ws = wb['Динамика оттока']
s_r, s_t = styles(ws, 4), styles(ws, 11)
for r in range(4, 14):
    for k in range(1, 18): ws.cell(r, k).value = None
mov = collections.defaultdict(float); cnt = collections.defaultdict(int)
for x in mr['moves']: mov[(x['prod'], x['m'], x['et'])] += x['amt']; cnt[(x['prod'], x['m'], x['et'])] += x['n']
ET = ('0_creation', '1_expansion', '15_contraction', '2_churn')
run = {p: c['MRR'][p] - sum(mov[(p, m, e)] for m in MON for e in ET) for p in ('POS', 'SFA')}
r = 4; DYN = {}
for m in MON:
    v = ['%s 2026' % RUS[m]]
    for p, L in (('POS', 'BCDF'), ('SFA', 'HIJL')):
        n, loss = cnt[(p, m, '2_churn')], -mov[(p, m, '2_churn')]
        v += [n, loss, d['base_n'][m][p], '=%s%d/%s%d' % (L[0], r, L[2], r), run[p], '=%s%d/%s%d' % (L[1], r, L[3], r)]
        DYN['%s|%s' % (p, m)] = {'n': n, 'loss': loss, 'base_n': d['base_n'][m][p], 'base_mrr': run[p], 'new': cnt[(p, m, '0_creation')]}
        run[p] += sum(mov[(p, m, e)] for e in ET)
    v += [cnt[('POS', m, '0_creation')], cnt[('SFA', m, '0_creation')], '=N%d-B%d' % (r, r), '=O%d-H%d' % (r, r)]
    put(ws, r, v, s_r); r += 1
put(ws, r, ['ИТОГО %s 2026' % PERIOD, '=SUM(B4:B%d)' % (r - 1), '=SUM(C4:C%d)' % (r - 1), None, None, None, None, '=SUM(H4:H%d)' % (r - 1), '=SUM(I4:I%d)' % (r - 1),
            None, None, None, None] + ['=SUM(%s4:%s%d)' % (k, k, r - 1) for k in 'NOPQ'], s_t)
ws['A1'] = 'Помесячная динамика оттока (%s)' % PERIOD_LONG
fn = OUTP('Отток_подписок_причины_%s_2026_%s.xlsx' % (PERIOD.replace('–', '-'), SUF))
wb.save(fn)
json.dump({'tot': tot, 'clients': {p: len(partners[p]) for p in partners}, 'reasons': {'%s|%s' % k: v for k, v in agg.items()},
           'cats': {'%s|%s' % k: v for k, v in ca.items()}, 'DYN': DYN, 'top': top}, open(J('res_churn'), 'w'), ensure_ascii=False)
print(fn, tot, 'топ:', [(t[1], t[2], t[3], int(t[0])) for t in top])
