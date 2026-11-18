# -*- coding: utf-8 -*-
"""Поступления денег по месяцам: по продуктам и по способам оплаты (Excel)."""
import json, collections, openpyxl
from style_lib import *
from cfg_cut import *
d = json.load(open(J('cash')))
def prod(r): return {'POS': 'POS', 'SFA': 'SFA', 'POS/SFA': 'POS/SFA'}.get(r['prod'], 'Без сервера / другое')
def meth(r):
    if r['meth'] in ('Payme', 'Click', 'Paynet'): return r['meth']
    if r['meth'] == 'Ручной': return 'Вручную (банк, наличные)'
    j = r['jr'] or 'журнал не указан'
    for x in ('Payme', 'Click', 'Paynet'):
        if x.lower() in j.lower(): return x
    return j + ' (способ не указан)'
PR = ['SFA', 'POS', 'POS/SFA', 'Без сервера / другое']
inb = [r for r in d if r['pt'] == 'inbound']; outb = [r for r in d if r['pt'] == 'outbound']
P = collections.defaultdict(float); Mt = collections.defaultdict(float); N = collections.defaultdict(int); C = collections.defaultdict(float)
for r in inb: P[(r['m'], prod(r))] += r['s']; Mt[(r['m'], meth(r))] += r['s']; N[(r['m'], meth(r))] += r['n']
for r in outb: C[r['m']] += -r['s']
PR = [p for p in PR if any(P[(m, p)] for m in MON)]
BASE = ['Payme', 'Click', 'Paynet', 'Вручную (банк, наличные)']
METH = [x for x in BASE if any(Mt[(m, x)] for m in MON)] + sorted({meth(r) for r in inb} - set(BASE), key=lambda x: -sum(Mt[(m, x)] for m in MON))
fx = sum(r['fx'] for r in d); fx_odoo = sum(r['s_odoo'] for r in d if r['cur'] != 'UZS'); fx_k = sum(r['s'] for r in d if r['cur'] != 'UZS')
wb = openpyxl.Workbook(); wb.remove(wb.active)
# ---------- лист 1 ----------
ws = wb.create_sheet('Поступления по продуктам')
title(ws, 'A1', 'Поступления денег по месяцам и продуктам, %s 2026 (на %s)' % (PERIOD, CUT_RU))
hd = ['Месяц'] + PR + ['Поступило всего', 'Комиссии Payme / Click / Paynet', 'Чистые поступления', 'Комиссия, % от поступлений']
header(ws, 3, hd)
n = len(PR); L = openpyxl.utils.get_column_letter
cT, cC, cN, cP = L(2 + n), L(3 + n), L(4 + n), L(5 + n)
r = 4
for m in MON:
    datarow(ws, r, ['%s 2026' % RUS[m]] + [P[(m, p)] for p in PR] + ['=SUM(B%d:%s%d)' % (r, L(1 + n), r), -C[m], '=%s%d+%s%d' % (cT, r, cC, r), '=IF(%s%d=0,0,-%s%d/%s%d)' % (cT, r, cC, r, cT, r)],
            fmts=[None] + [NUM] * (n + 3) + [PCT])
    r += 1
totalrow(ws, r, ['ИТОГО %s 2026' % PERIOD] + ['=SUM(%s4:%s%d)' % (L(c), L(c), r - 1) for c in range(2, 5 + n)] + ['=IF(%s%d=0,0,-%s%d/%s%d)' % (cT, r, cC, r, cT, r)],
         fmts=[None] + [NUM] * (n + 3) + [PCT])
tot = r
datarow(ws, r + 1, ['Доля продукта в поступлениях'] + ['=%s%d/$%s$%d' % (L(c), tot, cT, tot) for c in range(2, 2 + n)], fmts=[None] + [PCT] * n, blue_inputs=False)
notes(ws, r + 3, [
  'Источник: Odoo, account.payment — проведённые платежи с 01.02.2026 по %s по дате платежа (как в отчёте «cashflow» в Odoo). Выгрузка %s.' % (CUT_RU, EXPORT_RU),
  'Продукт — по серверу клиента в платеже (поле «server_product»). «Без сервера / другое» — платежи без привязки к серверу.',
  'Платежи в долларах пересчитаны по курсу 12 000 сум/$: %s $ = %s сум (в Odoo по курсу на дату платежа — %s сум).' % (f'{fx:,.0f}'.replace(',', ' '), f'{fx_k:,.0f}'.replace(',', ' '), f'{fx_odoo:,.0f}'.replace(',', ' ')),
  'Комиссии — исходящие платежи Payme / Click / Paynet, которые Odoo создаёт на каждое поступление через эти системы. В отчёте «cashflow» в Odoo они попадают в колонку «Нет» со знаком минус.',
  'Поступления — это деньги по дате оплаты, а не выручка: годовые оплаты приходят разово, счета могут оплачиваться в другом месяце, часть денег приходит авансом.'])
widths(ws, {'A': 26, **{L(c): 17 for c in range(2, 6 + n)}})
ws.freeze_panes = 'B4'
# ---------- лист 2 ----------
ws = wb.create_sheet('По способам оплаты')
title(ws, 'A1', 'Поступления по способам оплаты, %s 2026 (на %s)' % (PERIOD, CUT_RU))
header(ws, 3, ['Месяц'] + METH + ['ИТОГО'])
k = len(METH); r = 4
for m in MON:
    datarow(ws, r, ['%s 2026' % RUS[m]] + [Mt[(m, x)] for x in METH] + ['=SUM(B%d:%s%d)' % (r, L(1 + k), r)], fmts=[None] + [NUM] * (k + 1)); r += 1
totalrow(ws, r, ['ИТОГО %s 2026' % PERIOD] + ['=SUM(%s4:%s%d)' % (L(c), L(c), r - 1) for c in range(2, 3 + k)], fmts=[None] + [NUM] * (k + 1)); t2 = r
datarow(ws, r + 1, ['Доля способа оплаты'] + ['=%s%d/$%s$%d' % (L(c), t2, L(2 + k), t2) for c in range(2, 2 + k)], fmts=[None] + [PCT] * k, blue_inputs=False)
datarow(ws, r + 2, ['Платежей, шт.'] + [sum(N[(m, x)] for m in MON) for x in METH] + ['=SUM(B%d:%s%d)' % (r + 2, L(1 + k), r + 2)], fmts=[None] + [NUM] * (k + 1), blue_inputs=False)
notes(ws, r + 4, ['Способ оплаты — поле «Payment type» в платеже. Где оно не заполнено, показан журнал, через который проведён платёж.',
                  'Суммы — без вычета комиссий; комиссии — на листе «Поступления по продуктам».'])
widths(ws, {'A': 26, **{L(c): 18 for c in range(2, 3 + k)}})
ws.freeze_panes = 'B4'
fn = OUTP('Поступления_денег_%s_2026_%s.xlsx' % (PERIOD.replace('–', '-'), SUF))
wb.save(fn)
res = {'P': {'%s|%s' % kk: v for kk, v in P.items()}, 'M': {'%s|%s' % kk: v for kk, v in Mt.items()}, 'C': C, 'PR': PR, 'METH': METH}
json.dump(res, open(J('res_cash'), 'w'), ensure_ascii=False)
print(fn, '| поступило', round(sum(P.values()) / 1e6, 1), 'комиссии', round(sum(C.values()) / 1e6, 1), '|', METH)
