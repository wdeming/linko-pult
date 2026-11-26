# -*- coding: utf-8 -*-
"""Клиенты с несколькими подписками, срез 30.09.2026."""
import json, collections
import openpyxl
from style_lib import *
from cfg_cut import *
from cls import PNAME

d = json.load(open(J('multi')))
multi, T = d['multi'], d['totals']
CUTNOTE = ("Источник: Odoo (odoo.linko.uz), sale.order. Срез базы на " + CUT_RU + " (выгрузка " + EXPORT_RU + "): "
           "подписки в статусах «в работе» и «на паузе» на эту дату. Клиент = partner_id подписки.")

def typ(p):
    if p['pos'] and p['sfa']: return 'Разные продукты (POS и SFA)'
    return 'Несколько подписок POS' if p['pos'] else 'Несколько подписок SFA'

groups = collections.defaultdict(lambda: [0, 0, 0.0])
for p in multi:
    t = typ(p); groups[t][0] += 1; groups[t][1] += p['pos'] + p['sfa']; groups[t][2] += p['mrr']
bycount = collections.defaultdict(lambda: [0, 0, 0.0])
for p in multi:
    k = p['pos'] + p['sfa']; bycount[k][0] += 1; bycount[k][1] += k; bycount[k][2] += p['mrr']

wb = openpyxl.Workbook(); wb.remove(wb.active)

ws = wb.create_sheet("Свод")
widths(ws, {'A':46,'B':20,'C':16,'D':14,'E':20,'F':24})
title(ws, "A1", "Клиенты с несколькими подписками (на " + CUT_RU + ")")
header(ws, 3, ['Показатель','Значение','Доля'])
rows = [('Всего клиентов в базе', T['clients'], None),
        ('Из них с двумя и более подписками', len(multi), '=B5/B4'),
        (None, None, None),
        ('Всего подписок в базе', T['subs'], None),
        ('Из них у клиентов с несколькими подписками', sum(p['pos'] + p['sfa'] for p in multi), '=B8/B7'),
        (None, None, None),
        ('MRR базы, сум/мес', T['mrr'], None),
        ('MRR клиентов с несколькими подписками, сум/мес', sum(p['mrr'] for p in multi), '=B11/B10'),
        ('Средний MRR такого клиента, сум/мес', '=B11/B5', None)]
r = 4
for a, b, c in rows:
    if a is not None:
        datarow(ws, r, [a, b, c], fill=GREY, fmts=[None, NUM, PCT])
    r += 1

section(ws, "A14", "Разрез: одинаковый продукт или разные")
header(ws, 15, ['Тип','Клиентов','Доля среди них','Подписок','MRR, сум/мес','Средний MRR на клиента'])
ORDER = [('Несколько подписок POS', BLUE), ('Несколько подписок SFA', GREEN), ('Разные продукты (POS и SFA)', YELLOW)]
r = 16
for t, fill in ORDER:
    v = groups[t]
    datarow(ws, r, [t, v[0], f"=B{r}/$B$19", v[1], v[2], f"=E{r}/B{r}"], fill=fill,
            fmts=[None, NUM, PCT, NUM, NUM, NUM])
    r += 1
totalrow(ws, 19, ['ИТОГО','=SUM(B16:B18)','=B19/B19','=SUM(D16:D18)','=SUM(E16:E18)','=E19/B19'],
         fmts=[None, NUM, PCT, NUM, NUM, NUM])

section(ws, "A21", "Сколько подписок у одного клиента")
header(ws, 22, ['Подписок у клиента','Клиентов','Подписок всего','MRR, сум/мес'])
r = 23
for k in sorted(bycount):
    v = bycount[k]
    datarow(ws, r, ['%d подписки' % k, v[0], v[1], v[2]], fill=GREY, fmts=[None, NUM, NUM, NUM])
    r += 1
totalrow(ws, r, ['ИТОГО', f"=SUM(B23:B{r-1})", f"=SUM(C23:C{r-1})", f"=SUM(D23:D{r-1})"],
         fmts=[None, NUM, NUM, NUM])
notes(ws, r + 2, [CUTNOTE,
    "Считаются подписки в статусах «в работе» и «на паузе». Пауза входит в базу — клиент не ушёл.",
    "«Разные продукты» — у клиента есть и POS-, и SFA-подписка. «Несколько подписок POS/SFA» — две и более подписки внутри одной линейки.",
    "MRR валютных подписок пересчитан в сумы по курсу 12 000 сум/$."])

ws = wb.create_sheet("Клиенты")
widths(ws, {'A':5,'B':38,'C':30,'D':13,'E':8,'F':8,'G':13,'H':11,'I':20,'J':46})
title(ws, "A1", "Клиенты с двумя и более подписками (на " + CUT_RU + ")")
header(ws, 3, ['№','Клиент','Тип','Подписок','POS','SFA','Действующих','На паузе','MRR всего, сум/мес','Номера подписок'])
r = 4
for i, p in enumerate(multi, 1):
    isboth = bool(p['pos'] and p['sfa'])
    nums = ", ".join("%s (%s%s)" % (s['name'], PNAME.get(s['prod'], s['prod']), ', пауза' if s['st'] == 'На паузе' else '') for s in p['subs'])
    datarow(ws, r, [i, p['name'], typ(p), p['pos'] + p['sfa'], p['pos'], p['sfa'],
                    sum(1 for s in p['subs'] if s['st'] == 'В работе' and s['prod'] in ('POS', 'SFA')),
                    sum(1 for s in p['subs'] if s['st'] == 'На паузе' and s['prod'] in ('POS', 'SFA')), p['mrr'], nums],
            fill=(YELLOW if isboth else GREEN), fmts=[NUM, None, None, NUM, NUM, NUM, NUM, NUM, NUM, None], bold=isboth)
    r += 1
totalrow(ws, r, [None,'ИТОГО',None,f"=SUM(D4:D{r-1})",f"=SUM(E4:E{r-1})",f"=SUM(F4:F{r-1})",
                 f"=SUM(G4:G{r-1})",f"=SUM(H4:H{r-1})",f"=SUM(I4:I{r-1})",None],
         fmts=[None,None,None,NUM,NUM,NUM,NUM,NUM,NUM,None])
ws.freeze_panes = "C4"
notes(ws, r + 2, [CUTNOTE, "Жёлтым — клиенты, у которых есть оба продукта. Сортировка: по числу подписок, затем по MRR."])

ws = wb.create_sheet("Подписки детально")
widths(ws, {'A':5,'B':38,'C':30,'D':14,'E':10,'F':14,'G':20})
title(ws, "A1", "Все подписки этих клиентов (на " + CUT_RU + ")")
header(ws, 3, ['№','Клиент','Тип клиента','Подписка','Продукт','Статус','MRR, сум/мес'])
r = 4; i = 1
for p in multi:
    for s in p['subs']:
        datarow(ws, r, [i, p['name'], typ(p), s['name'], PNAME.get(s['prod'], s['prod']), s['st'], s['mrr']],
                fill=(BLUE if s['prod'] == 'POS' else GREEN), fmts=[NUM, None, None, None, None, None, NUM])
        r += 1; i += 1
totalrow(ws, r, [None,'ИТОГО',None,None,None,None,f"=SUM(G4:G{r-1})"], fmts=[None,None,None,None,None,None,NUM])
ws.freeze_panes = "D4"
notes(ws, r + 2, [CUTNOTE])

OUT = OUTP("Клиенты_с_несколькими_подписками_%s.xlsx" % SUF)
wb.save(OUT)
print("saved", OUT)
print("клиентов", len(multi), "подписок", sum(p['pos']+p['sfa'] for p in multi), "MRR", int(sum(p['mrr'] for p in multi)))
for t in groups: print(" ", t, groups[t][0], groups[t][1], int(groups[t][2]))
