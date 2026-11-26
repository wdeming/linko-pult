# -*- coding: utf-8 -*-
"""Клиентская база по продуктам и сегментам + клиенты с несколькими подписками, срез 30.09.2026."""
import json, collections
import openpyxl
from style_lib import *
from cfg_cut import *
from cls import PNAME

CUTNOTE = ("Источник: Odoo (odoo.linko.uz), sale.order. Срез базы на " + CUT_RU + " (выгрузка " + EXPORT_RU + "): "
           "подписки в статусах «в работе» и «на паузе» на эту дату; MRR каждой подписки — нарастающий итог журнала изменений MRR "
           "Odoo (sale.order.log) по " + CUT_RU + ", как во всём комплекте. Клиент = partner_id подписки.")
_c = json.load(open(J('const'))); _jo = sum((_c.get('MRR_JOUT') or {}).values())
if abs(_jo) > 1:
    CUTNOTE += (" Сверка с MRR по журналу: в журнале есть ещё %s сум по подпискам вне базы на срез (ошибки журнала — двойное закрытие и т.п.), "
                "поэтому MRR по журналу на срез = итог этого файла %s %s сум." % ('{:,.0f}'.format(_jo).replace(',', ' '), '+' if _jo > 0 else '−', '{:,.0f}'.format(abs(_jo)).replace(',', ' ')))
CURNOTE = "MRR валютных подписок пересчитан в сумы по курсу 12 000 сум/$ (управленческий курс)."

base = json.load(open(J('base')))
lines = json.load(open(J('lines')))
byorder = {r['id']: r for r in base}

# статус на срез — из выгрузки (по истории статусов)
for r in base:
    r['st31'] = '4_paused' if r['st'] == '4_paused' else '3_progress'
    r['stru'] = 'На паузе' if r['st31'] == '4_paused' else 'В работе'

POS = [r for r in base if r['prod'] == 'POS']
SFA = [r for r in base if r['prod'] == 'SFA']
AI = [r for r in base if r['prod'] == 'AI']; PS = [r for r in base if r['prod'] == 'PS']   # допродукты SFA — отдельные подписки

part = collections.defaultdict(lambda: {'pos': [], 'sfa': [], 'add': [], 'name': ''})
for r in base:
    p = part[r['partner_id']]
    p['name'] = r['pname']
    p['pos' if r['prod'] == 'POS' else 'add' if r['prod'] in ('AI', 'PS') else 'sfa'].append(r)

def seg(p):
    if p['pos'] and (p['sfa'] or p['add']): return 'POS и SFA (оба продукта)'   # допродукты AI / Perfect store — линейка SFA
    return 'Только POS' if p['pos'] else 'Только SFA'

segs = collections.defaultdict(lambda: [0, 0, 0.0])   # клиентов, подписок, MRR
for pid, p in part.items():
    s = seg(p); subs = p['pos'] + p['sfa'] + p['add']
    segs[s][0] += 1; segs[s][1] += len(subs); segs[s][2] += sum(x['mrr'] for x in subs)

TOTAL_CLIENTS = len(part)
TOTAL_SUBS = len(base)
TOTAL_MRR = sum(r['mrr'] for r in base)

# ============================ ФАЙЛ 1 ============================
wb = openpyxl.Workbook(); wb.remove(wb.active)

ws = wb.create_sheet("Свод по сегментам")
widths(ws, {'A':40,'B':14,'C':18,'D':14,'E':18,'F':20,'G':22,'H':18})
title(ws, "A1", "Linktrade — клиентская база по продуктам и сегментам (срез " + CUT_RU + ")")
header(ws, 3, ['Сегмент','Клиентов','Доля клиентов','Подписок','MRR, сум/мес','Доля MRR',
               'Средний MRR на клиента','Подписок на клиента'])
order = [('Только POS', BLUE), ('Только SFA', GREEN), ('POS и SFA (оба продукта)', YELLOW)]
r = 4
for s, fill in order:
    v = segs[s]
    datarow(ws, r, [s, v[0], f"=B{r}/$B$7", v[1], v[2], f"=E{r}/$E$7", f"=E{r}/B{r}", f"=D{r}/B{r}"],
            fill=fill, fmts=[None, NUM, PCT, NUM, NUM, PCT, NUM, DEC2])
    r += 1
totalrow(ws, 7, ['ИТОГО','=SUM(B4:B6)','=B7/B7','=SUM(D4:D6)','=SUM(E4:E6)','=E7/E7','=E7/B7','=D7/B7'],
         fmts=[None, NUM, PCT, NUM, NUM, PCT, NUM, DEC2])
ws.freeze_panes = "A4"

section(ws, "A9", "Клиенты, купившие оба продукта")
header(ws, 10, ['Клиент','Подписок POS','MRR POS','Подписок SFA','MRR SFA','Всего MRR, сум/мес'])
both = sorted([p for p in part.values() if p['pos'] and (p['sfa'] or p['add'])],
              key=lambda p: -sum(x['mrr'] for x in p['pos'] + p['sfa'] + p['add']))
r = 11
for p in both:
    datarow(ws, r, [p['name'], len(p['pos']), sum(x['mrr'] for x in p['pos']),
                    len(p['sfa']) + len(p['add']), sum(x['mrr'] for x in p['sfa'] + p['add']), f"=C{r}+E{r}"],
            fill=YELLOW, fmts=[None, NUM, NUM, NUM, NUM, NUM])
    r += 1
totalrow(ws, r, ['ИТОГО', f"=SUM(B11:B{r-1})", f"=SUM(C11:C{r-1})", f"=SUM(D11:D{r-1})",
                 f"=SUM(E11:E{r-1})", f"=SUM(F11:F{r-1})"], fmts=[None, NUM, NUM, NUM, NUM, NUM])
notes(ws, r + 2, [CUTNOTE,
    "Сегмент клиента определяется по набору продуктов в его подписках: только POS, только SFA либо оба продукта одновременно. AI ассистент и Perfect store — допродукты линейки SFA (отдельные подписки на сервере SFA), в сегментах считаются как SFA.",
    CURNOTE,
    "Статус «в работе / на паузе» взят по истории статусов Odoo на конец дня " + CUT_RU + ". Подписки, закрытые за неуплату и возвращённые в работу после среза, учтены как действующие."])

# --- лист 2 ---
ws = wb.create_sheet("Продукты — активные и пауза")
widths(ws, {'A':16,'B':22,'C':20,'D':20,'E':20,'F':18,'G':20,'H':26,'I':22})
title(ws, "A1", "Подписки по продуктам: действующие и на паузе (на " + CUT_RU + ")")
header(ws, 3, ['Продукт','Действующих подписок','MRR действующих','Подписок на паузе','MRR на паузе',
               'ВСЕГО подписок','ВСЕГО MRR','Клиентов с этим продуктом','Доля паузы в подписках'])
def split(rows):
    a = [x for x in rows if x['st31'] == '3_progress']; p = [x for x in rows if x['st31'] == '4_paused']
    return len(a), sum(x['mrr'] for x in a), len(p), sum(x['mrr'] for x in p)
for i, (nm, rows, fill) in enumerate([('POS', POS, BLUE), ('SFA', SFA, GREEN), (PNAME['AI'], AI, GREEN), (PNAME['PS'], PS, GREEN)]):
    r = 4 + i
    a, am, pn, pm = split(rows)
    cl = len({x['partner_id'] for x in rows})
    datarow(ws, r, [nm, a, am, pn, pm, f"=B{r}+D{r}", f"=C{r}+E{r}", cl, f"=IF(F{r}>0,D{r}/F{r},0)"],
            fill=fill, fmts=[None, NUM, NUM, NUM, NUM, NUM, NUM, NUM, PCT])
totalrow(ws, 8, ['ИТОГО','=SUM(B4:B7)','=SUM(C4:C7)','=SUM(D4:D7)','=SUM(E4:E7)','=SUM(F4:F7)',
                 '=SUM(G4:G7)', TOTAL_CLIENTS, '=D8/F8'],
         fmts=[None, NUM, NUM, NUM, NUM, NUM, NUM, NUM, PCT])
ws.freeze_panes = "A4"
nb = len(both)
notes(ws, 10, [CUTNOTE,
    "Продукт — по полю «Тип подписки»: основная POS, основная SFA, AI ассистент, Perfect store. AI ассистент и Perfect store — отдельные подписки на сервере SFA (с 26.11.2026).",
    "«Клиентов с этим продуктом» — уникальные клиенты; %d клиента имеют оба продукта (POS и SFA) и посчитаны "
    "в нескольких строках, как и клиенты с допродуктами, поэтому сумма по строкам больше итога (%d)." % (nb, TOTAL_CLIENTS),
    CURNOTE])

# --- лист 3: клиенты с 2+ подписками ---
multi = sorted([p for p in part.values() if len(p['pos']) + len(p['sfa']) > 1],      # 2+ ОСНОВНЫХ подписки; допродукты не в счёт
               key=lambda p: (-(len(p['pos']) + len(p['sfa'])), -sum(x['mrr'] for x in p['pos'] + p['sfa'] + p['add'])))
ws = wb.create_sheet("Клиенты с 2+ подписками")
widths(ws, {'A':5,'B':38,'C':20,'D':13,'E':8,'F':8,'G':13,'H':11,'I':20,'J':46})
title(ws, "A1", "Клиенты, у которых больше одной подписки (на " + CUT_RU + ")")
header(ws, 3, ['№','Клиент','Сегмент','Подписок всего','POS','SFA','Действующих','На паузе',
               'MRR всего, сум/мес','Номера подписок'])
r = 4
for i, p in enumerate(multi, 1):
    subs = p['pos'] + p['sfa'] + p['add']
    s = seg(p)
    nums = ", ".join("%s (%s%s)" % (x['name'], x['prod'], ', пауза' if x['st31'] == '4_paused' else '')
                     for x in sorted(subs, key=lambda x: -x['mrr']))
    isboth = bool(p['pos'] and p['sfa'])
    datarow(ws, r, [i, p['name'], s, len(subs), len(p['pos']), len(p['sfa']) + len(p['add']),
                    sum(1 for x in subs if x['st31'] == '3_progress'),
                    sum(1 for x in subs if x['st31'] == '4_paused'),
                    sum(x['mrr'] for x in subs), nums],
            fill=(YELLOW if isboth else GREEN), fmts=[NUM, None, None, NUM, NUM, NUM, NUM, NUM, NUM, None],
            bold=isboth)
    r += 1
totalrow(ws, r, [None, 'ИТОГО', None, f"=SUM(D4:D{r-1})", f"=SUM(E4:E{r-1})", f"=SUM(F4:F{r-1})",
                 f"=SUM(G4:G{r-1})", f"=SUM(H4:H{r-1})", f"=SUM(I4:I{r-1})", None],
         fmts=[None, None, None, NUM, NUM, NUM, NUM, NUM, NUM, None])
ws.freeze_panes = "C4"
nsub_multi = sum(len(p['pos']) + len(p['sfa']) + len(p['add']) for p in multi)
notes(ws, r + 2, [CUTNOTE,
    "Всего %d клиентов держат %d подписок. У %d клиентов (%.1f%%) больше одной подписки — на них приходится %d подписок."
    % (TOTAL_CLIENTS, TOTAL_SUBS, len(multi), len(multi) / TOTAL_CLIENTS * 100, nsub_multi),
    "Жёлтым выделены клиенты, у которых есть и POS, и SFA. Сортировка: по числу подписок, затем по MRR.",
    "В список попадают клиенты с 2+ основными подписками (POS / SFA). Допродукты AI ассистент и Perfect store — отдельные подписки, они показаны в номерах подписок и в столбце SFA, но клиента в список сами не добавляют."])

# --- лист 4: cross-use ---
lby = collections.defaultdict(list)
for l in lines: lby[l['pid']].append(l)
def modstat(pids):
    subs = {l['oid'] for p in pids for l in lby.get(p, [])}
    cls = {byorder[o]['partner_id'] for o in subs}
    amt = sum(l['amt'] for p in pids for l in lby.get(p, []))
    return len(subs), len(cls), amt
MODULES = [
    ('SFA full', 'SFA', 'Базовая лицензия SFA', [25], GREEN),
    ('SFA part', 'SFA', 'Ограниченная лицензия SFA', [26], GREEN),
    ('POS Подписка', 'POS', 'Базовая подписка POS', [24, 32], BLUE),
    ('POS бизнес / POS interprice', 'POS', 'Расширенный тариф POS', [40, 43], BLUE),
    ('Производство · Доработка', 'SFA (+POS)', 'Дополнительные модули в подписке', [36, 38, 39], YELLOW),
    ('Perfect store', 'SFA', 'Допродукт SFA — отдельная подписка', [35], YELLOW),
    ('AI ассистент', 'SFA', 'Допродукт SFA — отдельная подписка', [44], YELLOW),
    ('SFA full USD / SFA part USD', 'SFA', 'Валютные лицензии', [41, 42], GREEN),
]
ws = wb.create_sheet("Cross-use по модулям")
widths(ws, {'A':44,'B':14,'C':18,'D':74,'E':18,'F':60,'G':26})
title(ws, "A1", "Использование модулей: сколько клиентов платит больше чем за один (на " + CUT_RU + ")")
header(ws, 3, ['Модуль','Линейка','Что это','Подписок','Клиентов','Доля от базы клиентов','Сумма строк подписок, сум'])
r = 4
for nm, ln, what, pids, fill in MODULES:
    ns, nc, amt = modstat(pids)
    datarow(ws, r, [nm, ln, what, ns, nc, f"=E{r}/{TOTAL_CLIENTS}",
                    amt if pids != [41, 42] else 'в валюте документа'],
            fill=fill, fmts=[None, None, None, NUM, NUM, PCT, NUM], wrap=True)
    r += 1
def clients_with(pids):
    return {byorder[l['oid']]['partner_id'] for p in pids for l in lby.get(p, [])}
def clients_sub_with_both(a, b):
    oa = {l['oid'] for p in a for l in lby.get(p, [])}
    ob = {l['oid'] for p in b for l in lby.get(p, [])}
    return {byorder[o]['partner_id'] for o in oa & ob}
c_both_lines = len(both)
c_fullpart = len(clients_sub_with_both([25], [26]))
c_addon = len(clients_sub_with_both([25, 26], [36, 38, 39]) | ({p['partner_id'] for p in AI + PS} & {p['partner_id'] for p in SFA}))   # модуль в подписке SFA или допродукт отдельной подпиской
c_posmix = len(clients_sub_with_both([24, 32], [40, 43]))
mrr_both = sum(x['mrr'] for p in both for x in p['pos'] + p['sfa'] + p['add'])
section(ws, "A%d" % (r + 1), "Пересечения модулей")
header(ws, r + 2, ['Что пересекается','Комбинация','Клиентов','Доля от базы','MRR, сум/мес','Комментарий'])
rows2 = [
    ('Две продуктовые линейки: POS и SFA', 'POS + SFA', c_both_lines, mrr_both,
     'Клиент платит и за POS, и за SFA. Настоящая кросс-продажа между продуктами.', YELLOW),
    ('Две лицензии SFA в одной подписке', 'SFA full + SFA part', c_fullpart, '—',
     'Полные и ограниченные лицензии в одном договоре. Это стандартная комплектация, а не апсейл.', GREEN),
    ('Базовый SFA + дополнительный модуль', 'SFA full/part + Производство / Доработка / Perfect store / AI ассистент', c_addon, '—',
     'Клиент докупил функциональность сверх базовой лицензии. Настоящий апсейл внутри SFA.', GREEN),
    ('Базовая подписка POS + расширенный тариф', 'POS Подписка + POS бизнес', c_posmix, '—',
     '«POS бизнес» продаётся вместо базовой подписки, а не в дополнение к ней.', BLUE),
]
rr = r + 3
for a, b, c, e, f, fill in rows2:
    datarow(ws, rr, [a, b, c, f"=C{rr}/{TOTAL_CLIENTS}", e, f], fill=fill,
            fmts=[None, None, NUM, PCT, NUM, None], wrap=True)
    rr += 1
section(ws, "A%d" % (rr + 1), "ИТОГО cross-use")
header(ws, rr + 2, ['Трактовка','Клиентов','Доля от базы %d' % TOTAL_CLIENTS, 'Что считаем'])
strict, mid, wide = c_both_lines, c_both_lines + c_addon, c_both_lines + c_addon + c_fullpart + len(clients_with([40, 43]))
rows3 = [
    ('Строгая: две продуктовые линейки', strict, 'Клиент платит и за POS, и за SFA.', YELLOW),
    ('Средняя: апсейл внутри линейки', mid,
     '%d клиентов POS+SFA плюс %d клиентов с дополнительными модулями SFA (Производство, Perfect store, доработки).' % (c_both_lines, c_addon), GREEN),
    ('Широкая: любые два платных модуля', wide,
     'Плюс %d клиентов с комбинацией SFA full + SFA part и %d клиентов на расширенном тарифе POS. '
     'Но full/part — это типы лицензий в одном договоре, а не разные модули.' % (c_fullpart, len(clients_with([40, 43]))), BLUE),
]
rr2 = rr + 3
for a, b, c, fill in rows3:
    datarow(ws, rr2, [a, b, f"=B{rr2}/{TOTAL_CLIENTS}", c], fill=fill, fmts=[None, NUM, PCT, None], wrap=True)
    rr2 += 1
notes(ws, rr2 + 1, [CUTNOTE,
    "Модуль = товарная позиция в строке подписки (sale.order.line.product_id) по действующим и приостановленным подпискам на " + CUT_RU + ".",
    "Пересечения посчитаны по строкам одной подписки: комбинация засчитывается, если оба модуля есть в одном договоре. Кросс-продажа POS+SFA считается по клиенту (у него две разные подписки).",
    "Разовые услуги (обучение POS, обучение SFA, обучение AI) в состав модулей не включены — это не подписочная выручка.",
    "Главный вывод по строгой трактовке: кросс-продажи между линейками почти нет — %d клиентов из %d." % (strict, TOTAL_CLIENTS)])

OUT = OUTP("Клиентская_база_продукты_сегменты_%s.xlsx" % SUF)
wb.save(OUT)
print("saved", OUT)
print("клиентов", TOTAL_CLIENTS, "подписок", TOTAL_SUBS, "MRR", int(TOTAL_MRR))
for s in segs: print(" ", s, segs[s][0], segs[s][1], int(segs[s][2]))
print("многоподписочных:", len(multi), "подписок у них:", nsub_multi)
print("cross:", "both", c_both_lines, "full+part", c_fullpart, "addon", c_addon, "posmix", c_posmix)
json.dump({'multi': [{'name': p['name'], 'pos': len(p['pos']), 'sfa': len(p['sfa']), 'add': len(p['add']),
                      'mrr': sum(x['mrr'] for x in p['pos'] + p['sfa'] + p['add']),
                      'subs': [{'name': x['name'], 'prod': x['prod'], 'st': x['stru'], 'mrr': x['mrr']}
                               for x in sorted(p['pos'] + p['sfa'] + p['add'], key=lambda y: -y['mrr'])]}
                     for p in multi],
           'totals': {'clients': TOTAL_CLIENTS, 'subs': TOTAL_SUBS, 'mrr': TOTAL_MRR}},
          open(J('multi'),'w'), ensure_ascii=False)
