# -*- coding: utf-8 -*-
"""Себестоимость поддержки, внедрения, сопровождения ключевых клиентов и инфраструктуры — ФАКТ месяца среза из реестра «Расходы» Odoo.
ФОТ — по отделам, указанным в строках расходов (оклад, KPI, прочие выплаты за период) + налоги ЗП, распределённые пропорционально выплатам отделов.
Хостинг — расходы статьи «Хостинг и инфраструктура» за период. Всё по месяцам периода (сумма делится по дням), доллары по 12 000."""
import json, collections, datetime, calendar
import openpyxl
from openpyxl.utils import get_column_letter as L
from style_lib import *
from cfg_cut import *
c = json.load(open(J('const'))); LN = json.load(open(J('exp_lines'), encoding='utf-8')); EX = json.load(open(J('exp'), encoding='utf-8'))
# SFA здесь — вся линейка SFA: основные подписки + допродукты AI ассистент и Perfect store (их обслуживают те же отделы и тот же хостинг)
_SF = [x for x in ('SFA', 'AI', 'PS') if x in c['MRR']]
MRR_POS, MRR_SFA = c['MRR']['POS'], sum(c['MRR'][x] for x in _SF)
N_POS_ACT, MRR_POS_ACT = c['SPLIT']['POS']['3_progress']; N_POS_PAUSE, MRR_POS_PAUSE = c['SPLIT']['POS']['4_paused']
N_SFA_ACT, MRR_SFA_ACT = [sum(c['SPLIT'][x]['3_progress'][i] for x in _SF) for i in (0, 1)]; N_SFA_PAUSE, MRR_SFA_PAUSE = [sum(c['SPLIT'][x]['4_paused'][i] for x in _SF) for i in (0, 1)]
D = datetime.date.fromisoformat
MS = [m for m in MON if m >= '2026-04']                      # помесячный учёт расходов в Odoo — с апреля 2026
def split(a, d1, d2):
    a1, b = D(d1), D(d2); n = (b - a1).days + 1; res = collections.defaultdict(float); cur = a1
    while cur <= b:
        me = min(b, cur.replace(day=calendar.monthrange(cur.year, cur.month)[1])); res[cur.strftime('%Y-%m')] += a * ((me - cur).days + 1) / n; cur = me + datetime.timedelta(days=1)
    return res
DEPS = [(4, 'Отдел тех. поддержки', 'SUP'), (13, 'Onboarding POS', 'ONP'), (10, 'Onboarding SFA', 'ONS'), (12, 'Отдел по работе с ключевыми клиентами', 'KA')]
PAYK = {'Оклад': 'Оклад', 'KPI': 'KPI', 'Материальная помощь': 'Прочие выплаты'}
F = collections.defaultdict(float); PAYALL = collections.defaultdict(float); TAX = collections.defaultdict(float)
for x in LN:
    if x['plan'] != 'Затраты на персонал': continue
    for m, v in split(x['amt'], x['d1'], x['d2']).items():
        if x['cat'] == 'Налоги ЗП': TAX[m] += v
        elif x['cat'] in PAYK and x['dep_id']:
            PAYALL[m] += v; F[(x['dep_id'], PAYK[x['cat']], m)] += v
for did, _, _ in DEPS:
    for m in MS:
        pay = sum(F[(did, k, m)] for k in ('Оклад', 'KPI', 'Прочие выплаты'))
        F[(did, 'Налоги ЗП (доля)', m)] = TAX[m] * pay / PAYALL[m] if PAYALL[m] else 0
HOST = collections.defaultdict(float); HOSTP = collections.defaultdict(float); HREC = []
for x in EX:
    if x['cat'] != 'Хостинг и инфраструктура': continue
    HOSTP[x['pay'][:7]] += x['amt'] if x['pay'] <= CUT else 0
    for mm, v in split(x['amt'], x['d1'], x['d2']).items(): HOST[mm] += v
    if x['d1'][:7] <= LAST_M and x['pay'] >= '2026-03-01': HREC.append(x)
STAFF = collections.defaultdict(list)
for e in sorted(json.load(open(J('staff'))), key=lambda e: (not e['head'], -e['sal'], e['name'])): STAFF[e['dep']].append(e)
NS = {n: len(STAFF[d]) for d, n, _ in DEPS}
wb = openpyxl.Workbook(); wb.remove(wb.active)
# ============ ФОТ по отделам ============
ws = wb.create_sheet('ФОТ по отделам')
title(ws, 'A1', 'ФОТ отделов себестоимости по месяцам — факт из реестра «Расходы» Odoo, за период (на %s)' % CUT_RU)
n = len(MS); cav = L(2 + n)
header(ws, 3, ['Отдел / статья'] + ['%s 2026' % RUS[m] for m in MS] + ['Среднее в месяц'])
r = 4; TOTROW = {}
for did, name, key in DEPS:
    section(ws, 'A%d' % r, name); r += 1; r0 = r
    for k in ('Оклад', 'KPI', 'Прочие выплаты', 'Налоги ЗП (доля)'):
        datarow(ws, r, ['    ' + k] + [round(F[(did, k, m)]) for m in MS] + ['=AVERAGE(B%d:%s%d)' % (r, L(1 + n), r)], fmts=[None] + [NUM] * (n + 1)); r += 1
    totalrow(ws, r, ['ИТОГО ' + name] + ['=SUM(%s%d:%s%d)' % (L(cc), r0, L(cc), r - 1) for cc in range(2, 3 + n)], fmts=[None] + [NUM] * (n + 1)); TOTROW[key] = r; r += 2
section(ws, 'A%d' % r, 'База распределения налогов'); r += 1
datarow(ws, r, ['Налоги ЗП всего (статья «Налоги ЗП»)'] + [round(TAX[m]) for m in MS], fmts=[None] + [NUM] * n); rt = r; r += 1
datarow(ws, r, ['Выплаты всех отделов (оклад, KPI, прочие)'] + [round(PAYALL[m]) for m in MS], fmts=[None] + [NUM] * n); rp = r; r += 1
datarow(ws, r, ['Налоги в % от выплат'] + ['=IF(%s%d=0,0,%s%d/%s%d)' % (L(cc), rp, L(cc), rt, L(cc), rp) for cc in range(2, 2 + n)], fmts=[None] + [PCT] * n, blue_inputs=False); r += 2
notes(ws, r, ['Источник: Odoo, реестр «Расходы» (x_custom_expenses), строки с отделом, статьи «Оклад», «KPI», «Материальная помощь». Выгрузка %s.' % EXPORT_RU,
              'Месяц — по периоду расхода («Период начало – конец»), сумма делится по дням. Оклад за месяц, выплаченный в следующем месяце, стоит в своём месяце.',
              'KPI — как фактически начислен за период (часто выплачивается за несколько месяцев разом), поэтому по месяцам скачет; столбец «Среднее» сглаживает.',
              'Налоги ЗП в реестре без отдела — распределены по отделам пропорционально их выплатам в том же месяце.',
              'Последний месяц может быть неполным: расходы, которые ещё не внесены в Odoo, сюда не попали.'])
widths(ws, {'A': 46, **{L(cc): 15 for cc in range(2, 3 + n)}}); ws.freeze_panes = 'B4'
FD = "'ФОТ по отделам'"; CM = L(1 + n)                      # колонка месяца среза
# ============ Хостинг ============
ws = wb.create_sheet('Хостинг')
title(ws, 'A1', 'Хостинг и инфраструктура по месяцам (на %s)' % CUT_RU)
header(ws, 3, ['Месяц', 'За период (в себестоимость)', 'По дате оплаты'])
r = 4
for m in MS: datarow(ws, r, ['%s 2026' % RUS[m], round(HOST[m]), round(HOSTP[m])], fmts=[None, NUM, NUM]); r += 1
totalrow(ws, r, ['Среднее в месяц', '=AVERAGE(B4:B%d)' % (r - 1), '=AVERAGE(C4:C%d)' % (r - 1)], fmts=[None, NUM, NUM]); HROW = {m: 4 + i for i, m in enumerate(MS)}
r += 2; header(ws, r, ['ID в Odoo', 'Дата оплаты', 'Период с', 'Период по', 'Сумма, сум', 'В долларах', 'Касса', 'Примечание']); r += 1
for x in sorted(HREC, key=lambda x: (x['d1'], x['id'])):
    datarow(ws, r, [x['id'], D(x['pay']), D(x['d1']), D(x['d2']), round(x['amt']), x['usd'] or None, x['kassa'], x['note']],
            fmts=[None, 'DD.MM.YYYY', 'DD.MM.YYYY', 'DD.MM.YYYY', NUM, '#,##0.00', None, None], blue_inputs=False); r += 1
notes(ws, r + 1, ['Источник: реестр «Расходы», статья «Хостинг и инфраструктура». Платёж за несколько месяцев делится по дням периода. Доллары — по 12 000 сум/$.'])
widths(ws, {'A': 16, 'B': 22, 'C': 16, 'D': 12, 'E': 16, 'F': 12, 'G': 20, 'H': 50}); ws.freeze_panes = 'A4'
# ============ Базы и допущения ============
ws = wb.create_sheet('Базы и допущения')
widths(ws, {'A': 52, 'B': 18, 'C': 10, 'D': 80}); title(ws, 'A1', 'Базы распределения и допущения')
header(ws, 3, ['Параметр', 'Значение', 'Ед.', 'Источник / комментарий'])
ROWS = [(7, 'Доля ФОТ Onboarding POS, относимая на ПОДДЕРЖКУ POS', 0.5, '%', 'ДОПУЩЕНИЕ: выделенного отдела поддержки POS нет, функцию выполняет Onboarding POS. Тайм-шита нет → 50/50.', PCT),
        (8, 'Курс USD → UZS', 12000, 'сум', 'Управленческий курс для всего комплекта: валютные подписки и расходы в долларах.', NUM),
        (10, 'MRR действующих подписок, POS', MRR_POS_ACT, 'сум/мес', 'Срез %s. %d подписок.' % (CUT_RU, N_POS_ACT), NUM),
        (11, 'MRR подписок на паузе, POS', MRR_POS_PAUSE, 'сум/мес', '%d подписок. Пауза считается действующей базой.' % N_POS_PAUSE, NUM),
        (12, 'MRR действующих подписок, SFA', MRR_SFA_ACT, 'сум/мес', 'Срез %s. %d подписок.' % (CUT_RU, N_SFA_ACT), NUM),
        (13, 'MRR подписок на паузе, SFA', MRR_SFA_PAUSE, 'сум/мес', '%d подписок.' % N_SFA_PAUSE, NUM)]
for rr, a, b, cc, dd, fmt in ROWS: datarow(ws, rr, [a, b, cc, dd], fill=GREY, fmts=[None, fmt, None, None], wrap=True)
for rr, a, b, cc, dd, fmt in [(15, 'MRR SFA (с паузой)', '=B12+B13', 'сум/мес', '', NUM), (16, 'MRR POS (с паузой)', '=B10+B11', 'сум/мес', '', NUM), (17, 'MRR всего', '=B16+B15', 'сум/мес', '', NUM),
                              (18, 'Доля POS в MRR — база деления хостинга', '=B16/B17', '%', 'В расходах на хостинг признака продукта нет — делим по MRR.', PCT), (19, 'Доля SFA в MRR', '=B15/B17', '%', '', PCT),
                              (21, 'Подписок в базе, POS (действующие + пауза)', N_POS_ACT + N_POS_PAUSE, 'шт.', '', NUM), (22, 'Подписок в базе, SFA (действующие + пауза)', N_SFA_ACT + N_SFA_PAUSE, 'шт.', '', NUM)]:
    datarow(ws, rr, [a, b, cc, dd], fill=GREY, fmts=[None, fmt, None, None], wrap=True)
BA = "'Базы и допущения'"; HS = "'Хостинг'"
# ============ Свод ============
ws = wb.create_sheet('Свод себестоимости')
widths(ws, {'A': 24, 'B': 52, 'C': 50, 'D': 19, 'E': 19, 'F': 19})
title(ws, 'A1', 'Linktrade — себестоимость за %s 2026: поддержка, внедрение, ключевые клиенты, инфраструктура (факт, на %s)' % (LAST_RU_L, CUT_RU))
header(ws, 3, ['Блок', 'Статья затрат', 'База отнесения на продукт', 'ВСЕГО, сум/мес', 'POS, сум/мес', 'SFA, сум/мес'])
ref = lambda key: "=%s!$%s$%d" % (FD, CM, TOTROW[key]); fm = [None, None, None, NUM, NUM, NUM]
section(ws, 'A4', 'ПОДДЕРЖКА')
datarow(ws, 5, [None, 'ФОТ отдела технической поддержки (%d чел. в штате)' % NS['Отдел тех. поддержки'], '100% на SFA', ref('SUP'), '=0', '=D5'], fill=GREY, fmts=fm, wrap=True)
datarow(ws, 6, [None, 'Функция поддержки внутри Onboarding POS', 'Доля ФОТ Onboarding POS (лист «Базы и допущения»)', ref('ONP') + '*%s!$B$7' % BA, '=D6', '=0'], fill=GREY, fmts=fm, wrap=True)
totalrow(ws, 7, [None, 'ИТОГО поддержка', None, '=D5+D6', '=E5+E6', '=F5+F6'], fmts=fm)
section(ws, 'A8', 'ВНЕДРЕНИЕ')
datarow(ws, 9, [None, 'ФОТ Onboarding SFA (%d чел. в штате)' % NS['Onboarding SFA'], '100% на SFA', ref('ONS'), '=0', '=D9'], fill=GREY, fmts=fm, wrap=True)
datarow(ws, 10, [None, 'ФОТ Onboarding POS, часть на внедрение (%d чел. в штате)' % NS['Onboarding POS'], 'Остаток ФОТ Onboarding POS', ref('ONP') + '*(1-%s!$B$7)' % BA, '=D10', '=0'], fill=GREY, fmts=fm, wrap=True)
totalrow(ws, 11, [None, 'ИТОГО внедрение', None, '=D9+D10', '=E9+E10', '=F9+F10'], fmts=fm)
section(ws, 'A12', 'СОПРОВОЖДЕНИЕ КЛЮЧЕВЫХ КЛИЕНТОВ')
datarow(ws, 13, [None, 'ФОТ отдела по работе с ключевыми клиентами (%d чел. в штате)' % NS['Отдел по работе с ключевыми клиентами'], '100% на SFA', ref('KA'), '=0', '=D13'], fill=GREY, fmts=fm, wrap=True)
totalrow(ws, 14, [None, 'ИТОГО сопровождение ключевых клиентов', None, '=D13', '=E13', '=F13'], fmts=fm)
section(ws, 'A15', 'ОБСЛУЖИВАНИЕ ИНФРАСТРУКТУРЫ')
datarow(ws, 16, [None, 'Хостинг и инфраструктура за %s' % LAST_RU_L, 'Пропорционально MRR: POS / SFA', '=%s!$B$%d' % (HS, HROW[LAST_M]), '=D16*%s!$B$18' % BA, '=D16*%s!$B$19' % BA], fill=GREY, fmts=fm, wrap=True)
totalrow(ws, 17, [None, 'ИТОГО инфраструктура', None, '=D16', '=E16', '=F16'], fmts=fm)
totalrow(ws, 18, [None, 'ИТОГО СЕБЕСТОИМОСТЬ', None, '=D7+D11+D14+D17', '=E7+E11+E14+E17', '=F7+F11+F14+F17'], fmts=fm)
section(ws, 'A20', 'Справочно')
datarow(ws, 21, [None, 'MRR подписок в базе (действующие + пауза)', None, '=E21+F21', '=%s!$B$16' % BA, '=%s!$B$15' % BA], fill=GREY, fmts=fm)
datarow(ws, 22, [None, 'Себестоимость в % от MRR', None, '=D18/D21', '=E18/E21', '=F18/F21'], fill=GREY, fmts=[None, None, None, PCT, PCT, PCT])
datarow(ws, 23, [None, 'Подписок в базе, шт.', None, '=E23+F23', '=%s!$B$21' % BA, '=%s!$B$22' % BA], fill=GREY, fmts=fm)
datarow(ws, 24, [None, 'Себестоимость на 1 подписку, сум/мес', None, '=D18/D23', '=E18/E23', '=F18/F23'], fill=GREY, fmts=fm)
notes(ws, 27, ['Как считалось:',
  '  • Это ФАКТ за %s 2026 из реестра «Расходы» Odoo, а не расчёт от окладов в карточках сотрудников. Месяц — по периоду расхода: оклад за %s, выплаченный позже, входит.' % (LAST_RU_L, LAST_RU_L),
  '  • ФОТ отдела = оклад + KPI + прочие выплаты по строкам расходов с этим отделом + налоги ЗП (распределены по отделам пропорционально выплатам). Помесячно — лист «ФОТ по отделам».',
  '  • Поддержка — отдел тех. поддержки целиком на SFA; поддержку POS выполняет Onboarding POS, его ФОТ делится 50/50 между поддержкой и внедрением (параметр).',
  '  • Внедрение — Onboarding SFA на SFA, остаток Onboarding POS на POS. Ключевые клиенты — целиком на SFA (удержание, не привлечение).',
  '  • Хостинг — статья «Хостинг и инфраструктура» за период месяца, делится между продуктами пропорционально MRR (в расходах признака продукта нет).',
  '  • KPI начисляется неравномерно (иногда за несколько месяцев разом), поэтому себестоимость по месяцам может скакать — на листе «ФОТ по отделам» есть среднее.',
  '  • SFA — вся линейка SFA: основные подписки и допродукты AI ассистент и Perfect store (отдельные подписки на сервере SFA с 26.11.2026).',
  'Что НЕ включено: ФОТ разработки, продаж и маркетинга, обед, аутсорс, аренда офиса и административные расходы, SaaS-подписки, связь.'])
# ============ Сотрудники (справочно) ============
ws = wb.create_sheet('Сотрудники отделов')
title(ws, 'A1', 'Сотрудники отделов себестоимости на %s (справочно, из карточек сотрудников)' % CUT_RU)
header(ws, 3, ['Отдел', 'Сотрудник', 'Должность', 'Оклад в карточке, сум/мес']); r = 4
for did, name, _ in DEPS:
    for e in STAFF[did]: datarow(ws, r, [name, e['name'], e['job'], e['sal']], fmts=[None, None, None, NUM], blue_inputs=False); r += 1
widths(ws, {'A': 36, 'B': 32, 'C': 40, 'D': 18})
wb._sheets = [wb['Свод себестоимости'], wb['ФОТ по отделам'], wb['Хостинг'], wb['Базы и допущения'], wb['Сотрудники отделов']]
OUT = OUTP('Себестоимость_поддержки_внедрения_инфраструктуры_%s.xlsx' % SUF); wb.save(OUT)
# --- расчёт для пульта и юнит-экономики
tot = lambda did, m: sum(F[(did, k, m)] for k in ('Оклад', 'KPI', 'Прочие выплаты', 'Налоги ЗП (доля)'))
V = {key: tot(did, LAST_M) for did, _, key in DEPS}; hm = HOST[LAST_M]; sh = MRR_POS / (MRR_POS + MRR_SFA)
cost_pos = V['ONP'] + hm * sh; cost_sfa = V['SUP'] + V['ONS'] + V['KA'] + hm * (1 - sh)
json.dump({'COGS': {'POS': cost_pos, 'SFA': cost_sfa}, 'KA': V['KA'], 'SUP': V['SUP'], 'ONP': V['ONP'], 'ONS': V['ONS'], 'HOST': hm, 'HOST_AVG': hm,
           'HOST_M': {m: HOST.get(m, 0) for m in MS}, 'HOST_PAY_M': {m: HOSTP.get(m, 0) for m in MS}, 'NS': NS,
           'SAL': {name: sum(e['sal'] for e in STAFF[did]) for did, name, _ in DEPS},
           'FOT_M': {key: {m: tot(did, m) for m in MS} for did, _, key in DEPS}}, open(J('cost'), 'w'), ensure_ascii=False)
print(OUT); print({k: round(v / 1e6, 1) for k, v in V.items()}, 'хостинг', round(hm / 1e6, 1), '| ИТОГО', round((cost_pos + cost_sfa) / 1e6, 1), 'POS', round(cost_pos / 1e6, 1), 'SFA', round(cost_sfa / 1e6, 1))
for did, _, key in DEPS: print(key, [round(tot(did, m) / 1e6, 1) for m in MS])
print('хостинг', [round(HOST[m] / 1e6, 1) for m in MS])
