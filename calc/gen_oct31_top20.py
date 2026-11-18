# -*- coding: utf-8 -*-
"""Концентрация выручки и условия топ-20, срез 30.09.2026."""
import json, collections, datetime
import openpyxl
from style_lib import *
from cfg_cut import *

CUT = CUT_D
NOTE = ("Источник: Odoo (odoo.linko.uz), sale.order — подписки в статусах «в работе» и «на паузе» "
        "на " + CUT_RU + " (выгрузка " + EXPORT_RU + "). Валютные подписки пересчитаны по курсу 12 000 сум/$.")

base = json.load(open(J('base')))
lines = json.load(open(J('lines')))
meta = json.load(open(J('meta')))
lic = collections.defaultdict(float)
for l in lines:
    if l['pid'] in (25, 26, 41, 42, 24, 32, 40, 43):
        lic[l['oid']] += l['qty']

cl = collections.defaultdict(lambda: {'name': '', 'subs': [], 'mrr': 0.0})
for r in base:
    c = cl[r['partner_id']]
    c['name'] = r['pname']; c['subs'].append(r); c['mrr'] += r['mrr']
ranked = sorted(cl.values(), key=lambda c: -c['mrr'])
TOTAL_MRR = sum(c['mrr'] for c in ranked)
N = len(ranked)

wb = openpyxl.Workbook(); wb.remove(wb.active)

# --- лист 1 ---
ws = wb.create_sheet("Концентрация выручки")
widths(ws, {'A':26,'B':12,'C':15,'D':22,'E':13,'F':22})
title(ws, "A1", "Концентрация выручки по клиентам (на " + CUT_RU + ")")
header(ws, 3, ['Группа клиентов','Клиентов','Доля клиентов','MRR группы, сум/мес','Доля в MRR','Средний MRR клиента'])
cum = 0.0; acc = []
for i, c in enumerate(ranked, 1):
    cum += c['mrr']; acc.append(cum)
r = 4
for k in (1, 3, 5, 10, 20, 30, 50, 100):
    datarow(ws, r, ['Топ-%d' % k, k, f"=B{r}/{N}", acc[k-1], f"=D{r}/{TOTAL_MRR:.0f}", f"=D{r}/B{r}"],
            fill=GREY, fmts=[None, NUM, PCT, NUM, PCT, NUM])
    r += 1
rest = N - 100
datarow(ws, r, ['Остальные %d' % rest, rest, f"=B{r}/{N}", TOTAL_MRR - acc[99],
                f"=D{r}/{TOTAL_MRR:.0f}", f"=D{r}/B{r}"], fill=GREY, fmts=[None, NUM, PCT, NUM, PCT, NUM])
r += 1
totalrow(ws, r, ['ВСЯ БАЗА', N, 1, TOTAL_MRR, 1, f"=D{r}/B{r}"], fmts=[None, NUM, PCT, NUM, PCT, NUM])
ws.freeze_panes = "A4"
p80 = next(i for i, v in enumerate(acc, 1) if v >= 0.8 * TOTAL_MRR)
section(ws, "A%d" % (r + 2), "Как читать")
notes(ws, r + 3, [
    "Ни один клиент не даёт больше %.1f%% выручки — критической зависимости от одного контрагента нет." % (ranked[0]['mrr'] / TOTAL_MRR * 100),
    "Половину выручки набирают примерно %d клиентов (%.0f%% базы). Топ-100 клиентов (%.0f%% базы) дают %.0f%%."
      % (next(i for i, v in enumerate(acc, 1) if v >= 0.5 * TOTAL_MRR),
         next(i for i, v in enumerate(acc, 1) if v >= 0.5 * TOTAL_MRR) / N * 100, 100 / N * 100, acc[99] / TOTAL_MRR * 100),
    "Оставшиеся %d клиентов приносят %.0f%% выручки при среднем чеке около %.0f тыс. сум — это в основном POS."
      % (rest, (TOTAL_MRR - acc[99]) / TOTAL_MRR * 100, (TOTAL_MRR - acc[99]) / rest / 1000),
    "Порог 80%% выручки достигается на %d-м клиенте (точный расчёт по полной выгрузке всех %d клиентов)." % (p80, N),
    NOTE])

# --- лист 2: топ-20 ---
TOP = ranked[:20]
ws = wb.create_sheet("Топ-20 клиентов")
widths(ws, {'A':5,'B':34,'C':12,'D':11,'E':12,'F':18,'G':11,'H':15,'I':26,'J':15,'K':17,'L':14,'M':20})
title(ws, "A1", "Топ-20 клиентов по MRR (на " + CUT_RU + ")")
header(ws, 3, ['№','Клиент','Продукт','Подписок','Лицензий','MRR, сум/мес','Доля в MRR','Накопленная доля',
               'Планы','Первый контракт','Ближайший срок действия','Дней до истечения','Канал'])
r = 4; run = 0.0
for i, c in enumerate(TOP, 1):
    subs = c['subs']; run += c['mrr']
    prods = sorted({s['prod'] for s in subs})
    ends = [datetime.date.fromisoformat(s['end_date']) for s in subs if s['end_date']]
    firsts = [s['first_contract'] for s in subs if s['first_contract']]
    emin = min(ends) if ends else None
    days = (emin - CUT).days if emin else None
    plans = ", ".join(sorted({meta[str(s['id'])]['plan'] or '—' for s in subs}))
    teams = ", ".join(sorted({meta[str(s['id'])]['team'] or '—' for s in subs}))
    paused = any(s['st'] == '4_paused' for s in subs)
    fill = YELLOW
    if paused or (days is not None and days <= 14):
        fill = PatternFill('solid', fgColor='FFF8CBAD')
    elif days is not None and days <= 60:
        fill = PatternFill('solid', fgColor='FFFFF2CC')
    else:
        fill = GREEN if prods == ['SFA'] else BLUE
    datarow(ws, r, [i, c['name'], "/".join(prods), len(subs), sum(lic.get(s['id'], 0) for s in subs),
                    c['mrr'], f"=F{r}/{TOTAL_MRR:.0f}", run / TOTAL_MRR, plans,
                    min(firsts) if firsts else None, emin.isoformat() if emin else None, days, teams],
            fill=fill, fmts=[NUM, None, None, NUM, NUM, NUM, PCT, PCT, None, None, None, NUM, None])
    r += 1
totalrow(ws, r, [None, 'ИТОГО топ-20', None, f"=SUM(D4:D{r-1})", f"=SUM(E4:E{r-1})", f"=SUM(F4:F{r-1})",
                 f"=F{r}/{TOTAL_MRR:.0f}", None, None, None, None, None, None],
         fmts=[None, None, None, NUM, NUM, NUM, PCT, None, None, None, None, None, None])
ws.freeze_panes = "C4"
notes(ws, r + 2, [NOTE,
    "«Ближайший срок действия» — самая ранняя дата окончания среди подписок клиента; по ней Odoo автоматически закрывает подписку, если не продлить. «Дней до истечения» считается от " + CUT_RU + ". Сроки действия и даты следующего счёта — по состоянию на дату выгрузки.",
    "Оранжевым — клиенты, у которых срок истекал в ближайшие две недели от среза или подписка на паузе. Жёлтым — срок в ближайшие два месяца.",
    "«Лицензий» — суммарное количество рабочих мест SFA (full + part) или точек POS по всем подпискам клиента.",
    "Канал «Клиенты партнеров» означает, что клиент пришёл и обслуживается через партнёра — это дополнительный риск концентрации."])

# --- лист 3: условия ---
ws = wb.create_sheet("Условия и сроки")
widths(ws, {'A':5,'B':32,'C':12,'D':10,'E':16,'F':13,'G':18,'H':11,'I':15,'J':15,'K':14,'L':15,'M':20})
title(ws, "A1", "Условия подписок топ-20 клиентов (на " + CUT_RU + ")")
header(ws, 3, ['№','Клиент','Подписка','Продукт','Тарифный план','Статус','MRR, сум/мес','Лицензий',
               'Первый контракт','Срок действия','Дней до истечения','Следующий счёт','Канал'])
r = 4; i = 1
for c in TOP:
    for s in sorted(c['subs'], key=lambda x: -x['mrr']):
        m = meta[str(s['id'])]
        end = datetime.date.fromisoformat(s['end_date']) if s['end_date'] else None
        days = (end - CUT).days if end else None
        annual = (m['plan'] or '').startswith('12')
        paused = s['st'] == '4_paused'
        fill = PatternFill('solid', fgColor='FFF8CBAD') if (paused or (days is not None and days <= 14)) \
               else (BLUE if annual else GREEN)
        datarow(ws, r, [i, c['name'], s['name'], s['prod'], m['plan'],
                        'На паузе' if paused else 'В работе', s['mrr'], lic.get(s['id'], 0),
                        s['first_contract'], end.isoformat() if end else None, days, s['next_invoice'], m['team']],
                fill=fill, fmts=[NUM, None, None, None, None, None, NUM, NUM, None, None, NUM, None, None])
        r += 1; i += 1
totalrow(ws, r, [None, 'ИТОГО', None, None, None, None, f"=SUM(G4:G{r-1})", f"=SUM(H4:H{r-1})",
                 None, None, None, None, None],
         fmts=[None, None, None, None, None, None, NUM, NUM, None, None, None, None, None])
ws.freeze_panes = "D4"
notes(ws, r + 2, [NOTE,
    "ЧЕГО НЕТ В ODOO: сроков договорных обязательств, условий расторжения, штрафов, SLA, индексации цены, эксклюзивов. «Срок действия» здесь — это докуда оплачено, а не докуда клиент связан обязательством.",
    "Все подписки топ-20 идут по прайс-листу «По умолчанию UZS» без индивидуальных скидок: SFA full 120 000 сум за лицензию в месяц, SFA part 84 000. Годовой план — та же цена × 12, полугодовой × 6. Исключение — USD-прайс.",
    "Синим выделены годовые подписки (план «12 месяцев») — по ним клиент оплатил вперёд и до окончания срока не уйдёт.",
    "Оранжевым — истекало в ближайшие две недели от среза либо уже на паузе."])

OUT = OUTP("Концентрация_выручки_и_условия_топ20_%s.xlsx" % SUF)
wb.save(OUT)
print("saved", OUT)
print("клиентов", N, "MRR", int(TOTAL_MRR))
for k in (1,5,10,20,50,100): print(" топ-%d"%k, int(acc[k-1]), "%.1f%%"%(acc[k-1]/TOTAL_MRR*100))
print("порог 80%% на клиенте", p80)
risk=[(c['name'], int(c['mrr'])) for c in TOP
      if any(s['end_date'] and datetime.date.fromisoformat(s['end_date'])<=CUT_D+datetime.timedelta(days=14) for s in c['subs'])
      or any(s['st']=='4_paused' for s in c['subs'])]
print("под риском в топ-20:", len(risk), sum(x[1] for x in risk))
for x in risk: print("   ",x)
