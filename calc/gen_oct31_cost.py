# -*- coding: utf-8 -*-
"""Себестоимость поддержки, внедрения и инфраструктуры — на срез cfg_cut.CUT."""
import json, collections
import openpyxl
from style_lib import *

from cfg_cut import *
import datetime as _dt
c = json.load(open(J('const')))
MRR_POS, MRR_SFA = c['MRR']['POS'], c['MRR']['SFA']
N_POS_ACT, MRR_POS_ACT = c['SPLIT']['POS']['3_progress']; N_POS_PAUSE, MRR_POS_PAUSE = c['SPLIT']['POS']['4_paused']
N_SFA_ACT, MRR_SFA_ACT = c['SPLIT']['SFA']['3_progress']; N_SFA_PAUSE, MRR_SFA_PAUSE = c['SPLIT']['SFA']['4_paused']

def prov(note):
    n = note.lower()
    for k, v in (('netcup', 'NETCUP (серверы Ритм)'), ('cloudflare', 'Cloudflare'), ('hetzner', 'HETZNER'),
                 ('serve', 'Servercore'), ('proxima', 'Proxima'), ('exadot', 'Exadot'), ('новый сервер', 'Новый сервер'),
                 ('ритм', 'Ритм (сервера)'), ('ritm', 'Ритм (сервера)')):
        if k in n: return v
    return note[:30] or '—'
def month_of(d1, d2):
    """Месяц периода = месяц, на который приходится большая часть дней периода."""
    a = _dt.date.fromisoformat(d1); b = _dt.date.fromisoformat(d2)
    if b < a: b = a
    days = collections.Counter()
    d = a
    while d <= b:
        days[d.strftime('%Y-%m')] += 1; d += _dt.timedelta(days=1)
    return max(sorted(days), key=lambda m: days[m])
HOSTING = []
for h in json.load(open(J('hosting'))):
    usd = h['lsum'] if h['rate'] > 0 else 0
    val = h['lsum'] if h['rate'] == 0 else 0
    HOSTING.append((h['id'], h['pay'], h['d1'], h['d2'], prov(h['note']), usd, h['rate'], val, month_of(h['d1'], h['d2'])))
HOSTING = [h for h in HOSTING if h[8] <= LAST_M]
DEPN = {4: 'Отдел тех. поддержки', 13: 'Onboarding POS', 10: 'Onboarding SFA', 12: 'Отдел по работе с ключевыми клиентами'}
STAFF = {v: [] for v in DEPN.values()}
for e in sorted(json.load(open(J('staff'))), key=lambda e: (not e['head'], -e['sal'], e['name'])):
    STAFF[DEPN[e['dep']]].append((e['name'], e['job'], e['sal']))
NS = {k: len(v) for k, v in STAFF.items()}

wb = openpyxl.Workbook(); wb.remove(wb.active)

# ============ ФОТ детально ============
ws = wb.create_sheet("ФОТ детально")
widths(ws, {'A':5,'B':24,'C':30,'D':40,'E':17,'F':15,'G':17,'H':16,'I':20})
title(ws, "A1", "ФОТ поддержки, внедрения и сопровождения ключевых клиентов — полная стоимость сотрудников")
header(ws, 3, ['№','Отдел','Сотрудник','Должность','Оклад на руки, сум/мес','KPI, сум/мес',
               'Итого выплаты сотруднику','Налог с ФОТ, сум/мес','ПОЛНАЯ СТОИМОСТЬ, сум/мес'])
BA = "'Базы и допущения'"
r = 4; totals = {}
for dept, fill in (('Отдел тех. поддержки', GREEN), ('Onboarding POS', BLUE), ('Onboarding SFA', GREEN), ('Отдел по работе с ключевыми клиентами', GREEN)):
    start = r
    for i, (nm, job, sal) in enumerate(STAFF[dept], 1):
        datarow(ws, r, [i, dept, nm, job, sal,
                        f"=E{r}*{BA}!$B$4", f"=E{r}+F{r}",
                        f"=IF({BA}!$B$6=1,G{r}*{BA}!$B$5,G{r}/(1-{BA}!$B$5)-G{r})", f"=G{r}+H{r}"],
                fill=fill, fmts=[NUM, None, None, None, NUM, NUM, NUM, NUM, NUM])
        r += 1
    totalrow(ws, r, [None, None, None, 'ИТОГО %s (%d чел.)' % (dept, len(STAFF[dept])),
                     f"=SUM(E{start}:E{r-1})", f"=SUM(F{start}:F{r-1})", f"=SUM(G{start}:G{r-1})",
                     f"=SUM(H{start}:H{r-1})", f"=SUM(I{start}:I{r-1})"],
             fmts=[None, None, None, None, NUM, NUM, NUM, NUM, NUM])
    totals[dept] = r
    r += 2
notes(ws, r, ["Коэффициент перехода от оклада «на руки» к полной стоимости при текущих параметрах: (1 + KPI) × (1 + налог) = 1,20 × 1,075 = 1,29.",
              "Состав отделов — hr.employee на %s: действующие сотрудники и те, чья дата увольнения позже среза. Отдел по работе с ключевыми клиентами включён в себестоимость по методологическому решению менеджмента." % CUT_RU])
R_SUP, R_POS, R_SFA, R_KA = totals['Отдел тех. поддержки'], totals['Onboarding POS'], totals['Onboarding SFA'], totals['Отдел по работе с ключевыми клиентами']

# ============ Хостинг детально ============
ws = wb.create_sheet("Хостинг детально")
widths(ws, {'A':5,'B':9,'C':70,'D':12,'E':12,'F':13,'G':20,'H':13,'I':12,'J':20,'K':18})
title(ws, "A1", "Аренда серверов и хостинг — расшифровка")
header(ws, 3, ['№','ID Odoo','Дата платежа','Период с','Период по','Месяц периода','Поставщик',
               'Сумма, USD','Курс (rate)','Сумма в сумах (введена)','ИТОГО, сум'])
r = 4
for i, (oid, dt, d1, d2, prov, usd, rate, val, mper) in enumerate(HOSTING, 1):
    datarow(ws, r, [i, oid, dt, d1, d2, mper, prov, usd, rate, val, f"=IF(H{r}>0,H{r}*I{r},J{r})"],
            fill=GREY, fmts=[NUM, NUM, None, None, None, None, None, DEC2, NUM, NUM, NUM])
    r += 1
last = r - 1
totalrow(ws, r, [None, None, None, None, None, None, 'Итого за всё время', None, None, None,
                 f"=SUM(K4:K{last})"], fmts=[None]*10 + [NUM])
r += 2
section(ws, "A%d" % r, "Помесячно (по периоду, к которому относится платёж)")
r += 1
header(ws, r, ['Месяц периода','Сумма, сум','Комментарий'])
r += 1
first_m = r
AVGM = [m for m in MON if m >= '2026-04']
provs = collections.defaultdict(collections.Counter)
for h in HOSTING: provs[h[8]][h[4]] += 1
def cm(m):
    t = ' + '.join(k + (' ×%d' % v if v > 1 else '') for k, v in sorted(provs[m].items()))
    if m in ('2026-02', '2026-03'): return 'НЕПОЛНЫЙ месяц: только ' + t + ', проведена задним числом. В среднее не берётся.'
    return t
for m in MON:
    datarow(ws, r, [RUS[m] + ' 2026', f'=SUMIF($F$4:$F${last},"{m}",$K$4:$K${last})', cm(m)], fill=GREY, fmts=[None, NUM, None])
    r += 1
avg_row = r
totalrow(ws, r, ['СРЕДНЕМЕСЯЧНО, Апр–%s 2026 (%d мес)' % (LAST_SH, len(AVGM)), f"=AVERAGE(B{first_m+2}:B{r-1})",
                 'Период усреднения: все полные месяцы с апреля. Февраль и март исключены как неполные.'],
         fmts=[None, NUM, None])
notes(ws, r + 2, ['«Месяц периода» — месяц, на который приходится большая часть дней оплаченного периода: платёж за 26.10–25.11 отнесён к ноябрю и в срез не входит.',
                  'Источник: Odoo, x_custom_expenses, аналитика «Хостинг и инфраструктура» (выгрузка %s). Валютные платежи пересчитаны по курсу из поля rate каждой записи.' % EXPORT_RU])
AVG_REF = "'Хостинг детально'!$B$%d" % avg_row

# ============ Базы и допущения ============
ws = wb.create_sheet("Базы и допущения")
widths(ws, {'A':52,'B':18,'C':10,'D':80})
title(ws, "A1", "Базы распределения и допущения")
header(ws, 3, ['Параметр','Значение','Ед.','Источник / комментарий'])
ROWS = [
 (4,'KPI, % от оклада каждого сотрудника',0.2,'%','МЕТОДОЛОГИЧЕСКОЕ РЕШЕНИЕ МЕНЕДЖМЕНТА: KPI каждый месяц разный, закладываем 20% от оклада по всей выборке.',PCT),
 (5,'Налог с ФОТ, ставка',0.075,'%','Оклады в Odoo — это суммы «на руки», сверх них компания платит 7,5%.',PCT),
 (6,'Метод начисления налога: 1 = сверху, 0 = внутри начисленного',1,'1/0','1 — налог считается сверх выплат: стоимость = выплаты × (1 + ставка). 0 — gross-up: стоимость = выплаты / (1 − ставка). Разница около 0,6%. По умолчанию 1.',NUM),
 (7,'Доля ФОТ Onboarding POS, относимая на ПОДДЕРЖКУ POS',0.5,'%','ДОПУЩЕНИЕ: выделенного отдела поддержки POS нет, функцию выполняет Onboarding POS. Тайм-шита нет → 50/50.',PCT),
 (8,'Курс USD → UZS',12000,'сум','Управленческий курс, ориентир — res.currency.rate в Odoo на 01.08.2026. Применяется к валютным подпискам в MRR; в расходах курс берётся построчно из поля rate.',NUM),
 (10,'MRR действующих подписок, POS',MRR_POS_ACT,'сум/мес','Срез ' + CUT_RU + '. %d подписок.' % N_POS_ACT,NUM),
 (11,'MRR подписок на паузе, POS',MRR_POS_PAUSE,'сум/мес','%d подписок. По решению менеджмента пауза считается активной базой.' % N_POS_PAUSE,NUM),
 (12,'MRR действующих подписок, SFA',MRR_SFA_ACT,'сум/мес','Срез ' + CUT_RU + ', включая валютные подписки по курсу 12 000. %d подписок.' % N_SFA_ACT,NUM),
 (13,'MRR подписок на паузе, SFA',MRR_SFA_PAUSE,'сум/мес','%d подписок.' % N_SFA_PAUSE,NUM),
]
for r, a, b, c, dd, fmt in ROWS:
    datarow(ws, r, [a, b, c, dd], fill=GREY, fmts=[None, fmt, None, None], wrap=True)
CALC = [
 (15,'MRR SFA, всего в сумах (с паузой)','=B12+B13','сум/мес','Расчёт',NUM),
 (16,'MRR POS, всего в сумах (с паузой)','=B10+B11','сум/мес','Расчёт',NUM),
 (17,'MRR всего (POS + SFA)','=B16+B15','сум/мес','Расчёт. Сходится с отчётом «MRR_и_когортное_удержание_2026_%s.xlsx».' % SUF,NUM),
 (18,'Доля POS в MRR — база для деления хостинга','=B16/B17','%','В расходах на хостинг признака продукта нет — делим по выручке.',PCT),
 (19,'Доля SFA в MRR — база для деления хостинга','=B15/B17','%','',PCT),
 (21,'Действующих подписок POS',N_POS_ACT,'шт.','Статус «в работе» на конец дня ' + CUT_RU + ' по истории статусов Odoo (включая возвращённые из закрытия после среза)',NUM),
 (22,'Подписок на паузе POS',N_POS_PAUSE,'шт.','subscription_state = 4_paused',NUM),
 (23,'Действующих подписок SFA',N_SFA_ACT,'шт.','subscription_state = 3_progress',NUM),
 (24,'Подписок на паузе SFA',N_SFA_PAUSE,'шт.','subscription_state = 4_paused',NUM),
 (25,'Подписок в базе, POS (действующие + пауза)','=B21+B22','шт.','Расчёт',NUM),
 (26,'Подписок в базе, SFA (действующие + пауза)','=B23+B24','шт.','Расчёт',NUM),
]
for r, a, b, c, dd, fmt in CALC:
    datarow(ws, r, [a, b, c, dd], fill=GREY, fmts=[None, fmt, None, None], wrap=True)

# ============ Свод себестоимости ============
ws = wb.create_sheet("Свод себестоимости")
widths(ws, {'A':24,'B':52,'C':50,'D':19,'E':19,'F':19})
title(ws, "A1", "Linktrade — себестоимость поддержки, внедрения, сопровождения ключевых клиентов и инфраструктуры (на " + CUT_RU + ")")
header(ws, 3, ['Блок','Статья затрат','База отнесения на продукт','ВСЕГО, сум/мес','POS, сум/мес','SFA, сум/мес'])
FD = "'ФОТ детально'"
section(ws, "A4", "ПОДДЕРЖКА")
datarow(ws, 5, [None,'ФОТ отдела технической поддержки, %d чел. (полная стоимость)' % NS['Отдел тех. поддержки'],
                '100% на SFA — выделенная поддержка обслуживает SFA',
                f"={FD}!$I${R_SUP}", '=0', '=D5'], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM], wrap=True)
datarow(ws, 6, [None,'Функция поддержки внутри Onboarding POS',
                'Доля от полной стоимости Onboarding POS (параметр на листе «Базы и допущения»)',
                f"={FD}!$I${R_POS}*{BA}!$B$7", '=D6', '=0'], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM], wrap=True)
totalrow(ws, 7, [None,'ИТОГО поддержка',None,'=D5+D6','=E5+E6','=F5+F6'], fmts=[None,None,None,NUM,NUM,NUM])
section(ws, "A8", "ВНЕДРЕНИЕ")
datarow(ws, 9, [None,'ФОТ Onboarding SFA, %d чел. (полная стоимость)' % NS['Onboarding SFA'],'100% на SFA — напрямую',
                f"={FD}!$I${R_SFA}", '=0', '=D9'], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM], wrap=True)
datarow(ws, 10, [None,'ФОТ Onboarding POS, часть на внедрение','Остаток полной стоимости Onboarding POS после поддержки',
                 f"={FD}!$I${R_POS}*(1-{BA}!$B$7)", '=D10', '=0'], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM], wrap=True)
totalrow(ws, 11, [None,'ИТОГО внедрение',None,'=D9+D10','=E9+E10','=F9+F10'], fmts=[None,None,None,NUM,NUM,NUM])
section(ws, "A12", "СОПРОВОЖДЕНИЕ КЛЮЧЕВЫХ КЛИЕНТОВ")
datarow(ws, 13, [None,'ФОТ отдела по работе с ключевыми клиентами, %d чел. (полная стоимость)' % NS['Отдел по работе с ключевыми клиентами'],
                 '100% на SFA — ключевые клиенты (дистрибьюторы) обслуживаются по SFA',
                 f"={FD}!$I${R_KA}", '=0', '=D13'], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM], wrap=True)
totalrow(ws, 14, [None,'ИТОГО сопровождение ключевых клиентов',None,'=D13','=E13','=F13'], fmts=[None,None,None,NUM,NUM,NUM])
section(ws, "A15", "ОБСЛУЖИВАНИЕ ИНФРАСТРУКТУРЫ")
datarow(ws, 16, [None,'Аренда серверов и хостинг (среднее Апр–%s 2026)' % LAST_SH,'Пропорционально MRR продуктов: POS / SFA',
                 f"={AVG_REF}", f"=D16*{BA}!$B$18", f"=D16*{BA}!$B$19"], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM], wrap=True)
totalrow(ws, 17, [None,'ИТОГО инфраструктура',None,'=D16','=E16','=F16'], fmts=[None,None,None,NUM,NUM,NUM])
totalrow(ws, 18, [None,'ИТОГО СЕБЕСТОИМОСТЬ',None,'=D7+D11+D14+D17','=E7+E11+E14+E17','=F7+F11+F14+F17'], fmts=[None,None,None,NUM,NUM,NUM])
section(ws, "A20", "Справочно")
datarow(ws, 21, [None,'MRR подписок в базе (действующие + пауза), сум/мес',None,'=E21+F21',
                 f"={BA}!$B$16", f"={BA}!$B$15"], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM])
datarow(ws, 22, [None,'Себестоимость в % от MRR',None,'=D18/D21','=E18/E21','=F18/F21'], fill=GREY, fmts=[None,None,None,PCT,PCT,PCT])
datarow(ws, 23, [None,'Подписок в базе (действующие + пауза), шт.',None,'=E23+F23',
                 f"={BA}!$B$25", f"={BA}!$B$26"], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM])
datarow(ws, 24, [None,'Себестоимость на 1 подписку, сум/мес',None,'=D18/D23','=E18/E23','=F18/F23'], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM])
datarow(ws, 25, [None,'Справочно: себестоимость без ключевых клиентов',None,'=D18-D14','=E18-E14','=F18-F14'], fill=GREY, fmts=[None,None,None,NUM,NUM,NUM], wrap=True)
section(ws, "A28", "Как считалось:")
_K = 1.2 * 1.075
_sal = lambda d: sum(x[2] for x in STAFF[d])
_sh = MRR_POS / (MRR_POS + MRR_SFA) * 100
_f = lambda v: ('%.1f' % (v / 1e6)).replace('.', ',')
notes(ws, 29, [
 "  • ФОТ. Оклад в Odoo (x_studio_salary_1) — это сумма «на руки». Сверху начисляется KPI 20% от оклада и налог с ФОТ 7,5%. Полная стоимость = оклад × 1,20 × 1,075 = оклад × 1,29. Ставки и метод начисления налога меняются на листе «Базы и допущения».",
 "  • Поддержка — отдел технической поддержки (%d чел.) целиком на SFA. Поддержку POS выполняет Onboarding POS, поэтому его ФОТ делится между поддержкой и внедрением параметром (по умолчанию 50/50)." % NS['Отдел тех. поддержки'],
 "  • Внедрение — Onboarding SFA целиком на SFA, остаток Onboarding POS — на POS. Onboarding POS на %s — %d человек, ФОТ отдела %s млн «на руки»." % (CUT_RU, NS['Onboarding POS'], _f(_sal('Onboarding POS'))),
 "  • Сопровождение ключевых клиентов — отдел по работе с ключевыми клиентами (%d чел., %s млн «на руки», полная стоимость с KPI и налогом %s млн/мес) целиком на SFA: ключевые клиенты — дистрибьюторы, работающие на SFA. Это удержание и развитие действующих клиентов, а не привлечение, поэтому статья относится к себестоимости, а не к CAC (методологическое решение менеджмента)."
   % (NS['Отдел по работе с ключевыми клиентами'], _f(_sal('Отдел по работе с ключевыми клиентами')), _f(_sal('Отдел по работе с ключевыми клиентами') * _K)),
 "  • Обслуживание инфраструктуры — своих серверов нет, берём аренду: аналитика «Хостинг и инфраструктура» в x_custom_expenses, среднее за Апр–%s 2026 (%d полных месяцев). Валютные платежи пересчитаны по полю rate каждой записи. Февраль и март исключены как неполные." % (LAST_SH, len(AVGM)),
 "  • Хостинг разделён между продуктами пропорционально MRR подписок в базе на %s (POS ≈ %s%%, SFA ≈ %s%%) — в самих расходах признака продукта нет. Подписки на паузе (4_paused) включены в базу как действующие: клиент не ушёл, сервер под него продолжает работать."
   % (CUT_RU, ('%.1f' % _sh).replace('.', ','), ('%.1f' % (100 - _sh)).replace('.', ',')),
 "Что НЕ включено: ФОТ разработки, продаж и маркетинга (привлечение — в CAC), аренда офиса и административные расходы, SaaS-подписки и связь, амортизация, командировки и ГСМ выездных сотрудников."])
wb._sheets = [wb["Свод себестоимости"], wb["Базы и допущения"], wb["Хостинг детально"], wb["ФОТ детально"]]
OUT = OUTP("Себестоимость_поддержки_внедрения_инфраструктуры_%s.xlsx" % SUF)
wb.save(OUT)

# --- контрольный расчёт в питоне ---
K = 1.2 * 1.075
sup = sum(s[2] for s in STAFF['Отдел тех. поддержки']) * K
onp = sum(s[2] for s in STAFF['Onboarding POS']) * K
ons = sum(s[2] for s in STAFF['Onboarding SFA']) * K
ka = sum(s[2] for s in STAFF['Отдел по работе с ключевыми клиентами']) * K
host = {}
for oid, dt, d1, d2, prov, usd, rate, val, mper in HOSTING:
    host[mper] = host.get(mper, 0) + (usd * rate if usd > 0 else val)
avg = sum(host[m] for m in AVGM) / len(AVGM)
sh_pos = MRR_POS / (MRR_POS + MRR_SFA)
cost_pos = onp * 0.5 + onp * 0.5 + avg * sh_pos
cost_sfa = sup + ons + ka + avg * (1 - sh_pos)
print("saved", OUT)
print("ФОТ поддержка", int(sup), "| Onb POS", int(onp), "| Onb SFA", int(ons), "| KA", int(ka))
print("хостинг по месяцам:", {m: int(v) for m, v in sorted(host.items())})
print("среднее Апр–%s:" % LAST_SH, int(avg), "| доля POS %.2f%%" % (sh_pos * 100))
print("ИТОГО:", int(cost_pos + cost_sfa), "| POS", int(cost_pos), "| SFA", int(cost_sfa))
print("%% от MRR: POS %.1f%% SFA %.1f%% всего %.1f%%" % (cost_pos/MRR_POS*100, cost_sfa/MRR_SFA*100,
      (cost_pos+cost_sfa)/(MRR_POS+MRR_SFA)*100))
print("на подписку: POS", int(cost_pos/(N_POS_ACT+N_POS_PAUSE)), "SFA", int(cost_sfa/(N_SFA_ACT+N_SFA_PAUSE)))

json.dump({'COGS': {'POS': cost_pos, 'SFA': cost_sfa}, 'KA': ka, 'SUP': sup, 'ONP': onp, 'ONS': ons, 'HOST_AVG': avg,
           'HOST_M': {m: host.get(m, 0) for m in MON}, 'NS': NS,
           'SAL': {k: sum(x[2] for x in v) for k, v in STAFF.items()}}, open(J('cost'), 'w'), ensure_ascii=False)
