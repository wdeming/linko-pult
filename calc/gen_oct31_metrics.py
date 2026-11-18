# -*- coding: utf-8 -*-
"""Ключевые SaaS-метрики на срез cfg_cut.CUT — версия с учётом разовой выручки
(обучение, доработки, производство, Perfect store)."""
import json, collections
import openpyxl
from style_lib import *

from cfg_cut import *
ET = ('0_creation','1_expansion','15_contraction','2_churn')
PR = ('POS','SFA')
c = json.load(open(J('const'))); k = json.load(open(J('cost'))); mk = json.load(open(J('mkt')))
d = json.load(open(J('mrr')))
SRV, LIC, MRR, NEW = c['SRV'], c['LIC'], c['MRR'], c['NEW']
_tot = {p: sum(r['amt'] for r in d['moves'] if r['prod'] == p) for p in PR}
MRR0 = {p: MRR[p] - _tot[p] for p in PR}                      # MRR на 01.02.2026 обратным счётом
# себестоимость в карточке — без внедрения (оно учтено в CAC): поддержка + ключевые клиенты + хостинг
K = 1.2 * 1.075
COGS = {'POS': k['COGS']['POS'] - k['ONP'] * 0.5, 'SFA': k['COGS']['SFA'] - k['ONS']}
COGS_KA = {'POS': 0.0, 'SFA': k['KA']}
CYC_COH = c['CYC_COH']
# отток по количеству подписок: закрытий в месяц к базе на начало месяца (среднее за период)
_cnt = collections.defaultdict(int)
for r in d['events_n']: _cnt[(r['prod'], r['m'], r['et'])] = r['n']
def _churn_sub(p):
    n_end = SRV[p]; rates = []
    for m in reversed(MON):
        n_start = n_end - _cnt[(p, m, '0_creation')] + _cnt[(p, m, '2_churn')]
        rates.append(_cnt[(p, m, '2_churn')] / n_start if n_start else 0); n_end = n_start
    return sum(rates) / len(rates)
CHURN_SUB = {p: _churn_sub(p) for p in PR}
# ---- затраты на привлечение: факт Odoo, приведённый к полному периоду ----
# разметка ФОТ по отделам ведётся с мая 2026, аналитика расходов — с апреля; февраль–март (–апрель) проведены одной суммой.
FM = [m for m in MON if m >= '2026-05']; AM = [m for m in MON if m >= '2026-04']
def fot(key): return sum(mk['FOT_M'][key][m] for m in FM) / len(FM) * NM
FOT_SALES = {'POS': fot('sales_POS'), 'SFA': fot('sales_SFA')}
FOT_MKT = fot('marketing'); FOT_PARTNER = sum(mk['FOT_M']['partner'].values())
AD_SPEND   = sum(mk['POOL_M']['AD'][m] for m in AM) / len(AM) * NM
MKT_CONTR  = sum(mk['POOL_M']['MKT_CONTR'].values())
CRM_TOOLS  = sum(mk['POOL_M']['CRM'].values())
COMMISSION = sum(mk['POOL_M']['COMM'].values())
MKT_POOL   = FOT_MKT + AD_SPEND + MKT_CONTR + CRM_TOOLS
PARTNER_POOL = FOT_PARTNER + COMMISSION
DIG = {s: (mk['CH_LEADS'].get(s, 0), mk['CH_SUBS'][s]['POS'], mk['CH_SUBS'][s]['SFA'],
           mk['CH_SUBS'][s]['POS'] + mk['CH_SUBS'][s]['SFA']) for s in ('Инстаграм', 'Фейсбук', 'Сайт')}
DIG_LEADS = sum(v[0] for v in DIG.values())
DIG_SHARE = {'POS': sum(l*p/t for l, p, s_, t in DIG.values())/DIG_LEADS,
             'SFA': sum(l*s_/t for l, p, s_, t in DIG.values())/DIG_LEADS}
PARTNER_SUBS = dict(mk['CH_SUBS']['Партнер']); PARTNER_TOT = sum(PARTNER_SUBS.values())
CAC_SALES = {p: FOT_SALES[p] + MKT_POOL*DIG_SHARE[p] + PARTNER_POOL*PARTNER_SUBS[p]/PARTNER_TOT for p in PR}
CAC_ONB   = {'POS': fot('onb_POS'), 'SFA': fot('onb_SFA')}        # внедрение — факт, ВХОДИТ в CAC
PAID_SUBS = {p: sum(v[1 if p == 'POS' else 2] for v in DIG.values()) + PARTNER_SUBS[p] for p in PR}
ONCE, ONCE_M = c['ONCE'], c['ONCE_M']
OB = {int(a): b for a, b in c['ONCE_BY_PID'].items()}
EDU = mk['EDU']
mln = lambda v: ('%.1f' % (v / 1e6)).replace('.', ',')
mov = collections.defaultdict(float)
for r in d['moves']: mov[(r['prod'], r['m'], r['et'])] += r['amt']
series = {}
for p in PR:
    run = MRR0[p]; s = []
    for m in MON:
        st = run; run += sum(mov[(p, m, e)] for e in ET)
        s.append({'m': m, 'start': st, 'new': mov[(p,m,'0_creation')], 'exp': mov[(p,m,'1_expansion')],
                  'con': mov[(p,m,'15_contraction')], 'chu': mov[(p,m,'2_churn')], 'end': run})
    series[p] = s
avg = lambda p, f: sum(f(a) for a in series[p]) / NM
CH_CLOSED = {p: avg(p, lambda a: -a['chu']/a['start']) for p in PR}
CH_CONTR  = {p: avg(p, lambda a: -a['con']/a['start']) for p in PR}
EXPN      = {p: avg(p, lambda a: a['exp']/a['start']) for p in PR}

ONCE_MO = {p: ONCE[p]/NM for p in PR}                        # разовая выручка в среднем за месяц
REV_TOT = {p: MRR[p] + ONCE_MO[p] for p in PR}               # вся выручка месяца
GM_SUB  = {p: 1 - COGS[p]/MRR[p] for p in PR}                # маржа только по подписке
GM_TOT  = {p: 1 - COGS[p]/REV_TOT[p] for p in PR}            # маржа по всей выручке
ARPA    = {p: MRR[p]/SRV[p] for p in PR}
ARPU    = {p: REV_TOT[p]/SRV[p] for p in PR}
RPL     = {p: MRR[p]/LIC[p] for p in PR}
LPS     = {p: LIC[p]/SRV[p] for p in PR}
ONCE_NEW = {p: ONCE[p]/NEW[p] for p in PR}                   # разовая выручка на одну новую подписку
GP_CYCLE = {p: ARPA[p]*GM_TOT[p] for p in PR}                # маржа цикла
GP_ONCE  = {p: ONCE_NEW[p]*GM_TOT[p] for p in PR}            # разовая маржа при подключении
LIFE     = {p: 1/CH_CLOSED[p] for p in PR}
LTV      = {p: GP_ONCE[p] + GP_CYCLE[p]*LIFE[p] for p in PR}
LTV_SUB  = {p: ARPA[p]*GM_SUB[p]*LIFE[p] for p in PR}        # прежняя цифра, только подписка
LTV_COH  = {p: GP_ONCE[p] + GP_CYCLE[p]*CYC_COH[p] for p in PR}
CAC      = {p: (CAC_SALES[p]+CAC_ONB[p])/NEW[p] for p in PR}    # смешанный CAC, внедрение включено
CAC_IMPL = {p: CAC_SALES[p]/NEW[p] for p in PR}                  # справочно: без внедрения
CAC_PAID = {p: (MKT_POOL*DIG_SHARE[p] + PARTNER_POOL*PARTNER_SUBS[p]/PARTNER_TOT)/PAID_SUBS[p] for p in PR}

wb = openpyxl.Workbook(); wb.remove(wb.active)
ws = wb.create_sheet("Ключевые метрики")
widths(ws, {'A':44,'B':22,'C':22,'D':70})
title(ws, "A1", "Linktrade — ключевые SaaS-метрики на " + CUT_RU + " (с учётом разовой выручки)")
header(ws, 3, ['Показатель','POS','SFA','Как считается'])
r = 4
def row(name, v, fmt, how, fill=GREY, bold=False):
    global r
    datarow(ws, r, [name, v['POS'], v['SFA'], how], fill=fill,
            fmts=[None, fmt, fmt, None], bold=bold, wrap=True)
    r += 1; return r - 1

section(ws, "A%d" % r, "РАЗМЕР БИЗНЕСА"); r += 1
row('MRR (подписка), сум/мес', MRR, NUM, 'Нарастающий итог журнала изменений MRR в Odoo (создания, расширения, сокращения, отток) по ' + CUT_RU + ', как график «Изменение MRR»; USD по 12 000')
row('Разовая выручка, сум/мес в среднем', ONCE_MO, NUM,
    'POS: обучение + доработки POS. SFA: доработки, производство, Perfect store, обучение SFA и AI. '
    'Начислено за %s 2026 ÷ %d месяцев' % (PERIOD, NM), YELLOW)
row('ВСЕГО выручки, сум/мес', REV_TOT, NUM, 'MRR + разовая выручка', YELLOW, bold=True)
row('ARR (по подписке), сум/год', {p: MRR[p]*12 for p in PR}, NUM, 'MRR × 12')
row('Кол-во серверов (клиентских аккаунтов)', SRV, NUM, 'Уникальных x_studio_pos_server по подпискам базы')
row('Кол-во лицензий', LIC, NUM, 'Σ product_uom_qty по строкам подписок')
row('Лицензий на сервер', LPS, DEC2, 'Кол-во лицензий ÷ кол-во серверов')
r += 1

section(ws, "A%d" % r, "НА СЕРВЕР"); r += 1
row('ARPA — подписка на сервер, сум/мес', ARPA, NUM, 'MRR ÷ кол-во серверов')
row('ARPU — вся выручка на сервер, сум/мес', ARPU, NUM,
    'ВСЕГО выручки ÷ кол-во серверов. У POS выше ARPA в %s раза — обучение продаётся почти к каждому подключению' % ('%.2f' % (ARPU['POS']/ARPA['POS'])).replace('.', ','),
    YELLOW, bold=True)
row('ACV, сум/год на сервер', {p: ARPA[p]*12 for p in PR}, NUM, 'ARPA × 12 — стоимость подписочного контракта')
row('Валовая прибыль на сервер, сум/мес', {p: ARPU[p]*GM_TOT[p] for p in PR}, NUM, 'ARPU × Gross Margin')
r += 1

section(ws, "A%d" % r, "НА ЛИЦЕНЗИЮ — база для ценообразования"); r += 1
row('Выручка на лицензию, сум/мес', RPL, NUM, 'MRR ÷ кол-во лицензий. Прайс SFA: full 120 000, part 84 000')
row('Валовая прибыль на лицензию, сум/мес', {p: RPL[p]*GM_SUB[p] for p in PR}, NUM,
    'Выручка на лицензию × маржа по подписке')
r += 1

section(ws, "A%d" % r, "РАЗОВАЯ ВЫРУЧКА"); r += 1
row('Разовая выручка за %s 2026, сум' % PERIOD, ONCE, NUM,
    'Обучение POS %s млн + доработки POS %s млн | доработки SFA %s + Perfect store %s + производство %s + обучение SFA %s + обучение AI %s + AI ассистент %s'
    % tuple(mln(OB.get(x, 0)) for x in (30, 39, 38, 35, 36, 33, 37, 44)))
row('Новых подписок за период, шт', NEW, NUM, 'Новые подписки за %s 2026 (события создания в журнале MRR)' % PERIOD)
row('Разовая выручка на одну новую подписку, сум', ONCE_NEW, NUM,
    'Разовая выручка ÷ новые подписки — сколько клиент платит сверх подписки при подключении', YELLOW, bold=True)
row('Доля разовой в общей выручке', {p: ONCE_MO[p]/REV_TOT[p] for p in PR}, PCT, 'Разовая ÷ ВСЕГО выручки')
datarow(ws, r, ['Привязка к подключению', '%.0f%% суммы — новым клиентам, лаг %s дня' % (EDU['new_share']*100, ('%.1f' % EDU['lag']).replace('.', ',')),
                'производство и Perfect store — новым, доработки — и базе',
                'Проверено по дате счёта против даты первого контракта клиента. Обучение POS: %d клиентов, '
                'средний чек %s сум, покупают %.0f%% новых подключений'
                % (EDU['clients'], '{:,.0f}'.format(EDU['amt']/EDU['clients']).replace(',', ' '), EDU['clients']/NEW['POS']*100)],
        fill=GREY, fmts=[None]*4, wrap=True); r += 2

section(ws, "A%d" % r, "МАРЖА"); r += 1
row('Себестоимость, сум/мес', COGS, NUM,
    'Поддержка + сопровождение ключевых клиентов + хостинг — файл «Себестоимость_поддержки_внедрения_инфраструктуры_%s.xlsx». Внедрение из себестоимости исключено: оно учтено в CAC' % SUF)
row('в т.ч. отдел по работе с ключевыми клиентами', COGS_KA, NUM, '%d чел., %s млн «на руки» × 1,29, целиком на SFA' % (k['NS']['Отдел по работе с ключевыми клиентами'], mln(k['SAL']['Отдел по работе с ключевыми клиентами'])))
row('Gross Margin по подписке', GM_SUB, PCT, 'MRR минус повторяющаяся себестоимость ÷ MRR')
row('Gross Margin по всей выручке', GM_TOT, PCT,
    'ПРИНЯТА ЗА БАЗУ: та же себестоимость (поддержка + ключевые клиенты + хостинг) обслуживает и подписку, и разовые работы',
    YELLOW, bold=True)
r += 1

section(ws, "A%d" % r, "ОТТОК"); r += 1
row('Отток по деньгам (закрытые подписки), в мес', CH_CLOSED, PCT,
    'ПРИНЯТ ЗА БАЗУ ДЛЯ LTV: ушедший MRR закрытых подписок ÷ MRR на начало месяца, среднее за %s' % PERIOD, YELLOW, bold=True)
row('Справочно: отток по количеству подписок, в мес', CHURN_SUB, PCT, 'Закрытых подписок ÷ средняя база')
row('Справочно: сокращения действующих подписок, в мес', CH_CONTR, PCT, 'В отток НЕ включается (решение менеджмента)')
row('Справочно: расширения, в мес', EXPN, PCT, 'Прирост MRR действующих подписок ÷ MRR на начало')
row('Средний срок жизни сервера, мес', LIFE, DEC2, '1 ÷ отток по деньгам')
r += 1

section(ws, "A%d" % r, "LTV, CAC И ОКУПАЕМОСТЬ"); r += 1
row('Разовая маржа при подключении, сум', GP_ONCE, NUM, 'Разовая выручка на новую подписку × Gross Margin — приходит один раз')
row('Маржа цикла подписки, сум/мес', GP_CYCLE, NUM, 'ARPA × Gross Margin — приходит каждый месяц жизни')
r_ltv = row('LTV, сум', LTV, NUM,
            'Разовая маржа при подключении + маржа цикла × срок жизни', YELLOW, bold=True)
row('Справочно: LTV только по подписке', LTV_SUB, NUM, 'Прежний расчёт, без разовой выручки')
row('Справочно: LTV по когортной кривой', LTV_COH, NUM,
    'Срок жизни взят по когортам новых клиентов %s 2026, а не 1 ÷ отток; хвост кривой — по среднему темпу убыли последних трёх месяцев' % PERIOD)
row('Затраты на привлечение за %s 2026, сум' % PERIOD, CAC_SALES, NUM,
    'Фактические выплаты отделам продаж и маркетинга + реклама + подрядчики маркетинга + CRM + партнёрский отдел + комиссии партнёрам. '
    'Маркетинг разнесён по продуктам через digital-лиды, партнёрские затраты — по подпискам канала «Партнёр». '
    'ФОТ и реклама приведены к %d месяцам по среднемесячному факту (см. примечание)' % NM)
row('Затраты на внедрение за %s 2026, сум' % PERIOD, CAC_ONB, NUM,
    'Фактические выплаты отделам Onboarding POS и Onboarding SFA, приведены к %d месяцам. Входят в CAC' % NM)
r_cac = row('CAC на новый сервер (смешанный, с внедрением), сум', CAC, NUM,
            '(Затраты на привлечение + затраты на внедрение) ÷ новые подписки. Внедрение входит в CAC и исключено из себестоимости', YELLOW, bold=True)
row('Справочно: CAC платных каналов, сум', CAC_PAID, NUM,
    'Маркетинг + партнёрские затраты ÷ подписки из Инстаграма, Фейсбука, сайта и партнёров (без ФОТ продаж) — для решений по рекламному бюджету')
row('Справочно: CAC без затрат на внедрение, сум', CAC_IMPL, NUM, 'Затраты на привлечение ÷ новые подписки — если относить внедрение на себестоимость')
row('LTV : CAC', {p: LTV[p]/CAC[p] for p in PR}, DEC2, 'Ориентир здорового SaaS — от 3', YELLOW, bold=True)
row('Справочно: LTV : CAC без разовой выручки', {p: LTV_SUB[p]/CAC[p] for p in PR}, DEC2, 'Только подписочная маржа')
row('Справочно: LTV : CAC по когорте', {p: LTV_COH[p]/CAC[p] for p in PR}, DEC2, 'Консервативный вариант')
row('Payback, мес', {p: (CAC[p]-GP_ONCE[p])/GP_CYCLE[p] for p in PR}, DEC2,
    '(CAC − разовая маржа при подключении) ÷ маржа цикла: разовая выручка гасит часть CAC сразу', YELLOW)
row('Справочно: Payback без разовой выручки, мес', {p: CAC[p]/(ARPA[p]*GM_SUB[p]) for p in PR}, DEC2, 'Прежний расчёт')
r += 1

aug = {p: series[p][-1] for p in PR}; jul = {p: series[p][-2] for p in PR}
g = lambda a: (a['start']+a['chu']+a['con'])/a['start']
n = lambda a: (a['start']+a['chu']+a['con']+a['exp'])/a['start']
section(ws, "A%d" % r, "УДЕРЖАНИЕ ВЫРУЧКИ, %s 2026" % LAST_RU.upper()); r += 1
row('GRR (%s)' % LAST_RU_L, {p: g(aug[p]) for p in PR}, PCT, '(MRR на начало − отток − сокращения) ÷ MRR на начало')
row('NRR (%s)' % LAST_RU_L, {p: n(aug[p]) for p in PR}, PCT, 'То же плюс расширения')
r += 1
section(ws, "A%d" % r, "РОСТ"); r += 1
row('MRR на %s' % PREV_END_RU, {p: jul[p]['end'] for p in PR}, NUM, '')
row('MRR на %s' % CUT_RU, MRR, NUM, '')
row('MoM (Growth Rate)', {p: MRR[p]/jul[p]['end']-1 for p in PR}, PCT, 'Прирост MRR к предыдущему месяцу', YELLOW)
r += 1
notes(ws, r, [
 "СОСТАВ СЕБЕСТОИМОСТИ И CAC (методологические решения менеджмента). Себестоимость = поддержка + сопровождение ключевых клиентов "
 "(%d чел., %s млн/мес с KPI и налогом, целиком на SFA) + хостинг. Внедрение (Onboarding) относится к привлечению и входит в CAC, из себестоимости исключено. "
 "CAC = (расходы на маркетинг и продажи + внедрение) за период ÷ новые подписки за тот же период. Затраты на привлечение за %s: POS %s млн, SFA %s млн; "
 "на внедрение: POS %s млн, SFA %s млн."
 % (k['NS']['Отдел по работе с ключевыми клиентами'], mln(k['KA']), PERIOD, mln(CAC_SALES['POS']), mln(CAC_SALES['SFA']), mln(CAC_ONB['POS']), mln(CAC_ONB['SFA'])),
 "ПРИВЕДЕНИЕ РАСХОДОВ К ПЕРИОДУ. Разметка выплат по отделам в Odoo ведётся с мая 2026, аналитика расходов — с апреля; расходы февраля–марта проведены одной суммой без разбивки. "
 "Поэтому ФОТ продаж, маркетинга и внедрения взят как среднемесячный факт за май–%s (%d мес.) × %d мес. периода, реклама — среднемесячный факт за апрель–%s × %d мес. "
 "Подрядчики маркетинга, CRM и комиссии партнёрам — по факту. В версиях до 31.10.2026 расходы за неполное число месяцев делились на новые подписки всего периода, из-за чего CAC был занижен."
 % (LAST_RU_L, len(FM), NM, LAST_RU_L, NM),
 "ДВА CAC. Смешанный CAC (все каналы, включая территорию и рекомендации) — для оценки юнит-экономики. CAC платных каналов — "
 "только маркетинг и партнёрские затраты на подписки из платных каналов — для решений о рекламном бюджете. "
 "CAC считается по продукту отдельно: клиент POS (магазин) и клиент SFA (дистрибьютор) — разные сегменты, смешанная цифра для управления бесполезна.",
 "РАЗОВАЯ ВЫРУЧКА: обучение при подключении, доработки, производство, Perfect store, AI ассистент. Разовые продажи дают %.0f%% выручки POS и %.0f%% выручки SFA. "
 "В LTV разовая выручка входит ОДИН РАЗ, а не каждый месяц: %.0f%% суммы обучения POS выставлено новым клиентам, "
 "средний разрыв между датой счёта и датой первого контракта — %s дня. Это платёж при подключении, поэтому он прибавляется к LTV разово и вычитается из CAC при расчёте окупаемости."
 % (ONCE_MO['POS']/REV_TOT['POS']*100, ONCE_MO['SFA']/REV_TOT['SFA']*100, EDU['new_share']*100, ('%.1f' % EDU['lag']).replace('.', ',')),
 "СЕБЕСТОИМОСТЬ И РАЗОВАЯ ВЫРУЧКА. Вся выручка (подписка + разовая) сопоставляется с одной и той же суммой затрат. Маржа по подписке и по всей выручке: "
 "POS %.1f%% / %.1f%%, SFA %.1f%% / %.1f%%."
 % (GM_SUB['POS']*100, GM_TOT['POS']*100, GM_SUB['SFA']*100, GM_TOT['SFA']*100),
 "ОГОВОРКА ПО SFA: доработки и производство делает отдел разработки, чей ФОТ в себестоимость НЕ включён (он в OPEX). "
 "Поэтому маржа SFA по всей выручке немного завышена.",
 "ОТТОК И СРОК ИЗМЕРЕНИЯ. Подписки, закрытые за неуплату и возвращённые в работу после оплаты, в Odoo из журнала оттока исчезают. Поэтому отток последнего месяца, "
 "измеренный сразу после его окончания, завышен; цифры этого файла сняты %s и учитывают возвраты по эту дату." % EXPORT_RU,
 "ARPU выше ARPA, потому что клиент платит не только за подписку. У POS разовая выручка на одну новую "
 "подписку %.0f сум против ARPA %.0f сум в месяц." % (ONCE_NEW['POS'], ARPA['POS']),
 "Источник: Odoo (odoo.linko.uz, база linktrade), выгрузка %s, срез восстановлен на %s. "
 "Разовая выручка — account.move.line по проведённым счетам за 01.02–%s. Валютные подписки по курсу 12 000 сум/$ (управленческий курс)." % (EXPORT_RU, CUT_RU, CUT_RU),
])
ws.freeze_panes = "B4"

# ---------- лист 2 ----------
ws = wb.create_sheet("Помесячно 2026")
widths(ws, {'A':18,'B':18,'C':16,'D':16,'E':16,'F':16,'G':18,'H':12,'I':12,'J':12})
title(ws, "A1", "MRR, разовая выручка и удержание помесячно, %s" % PERIOD_LONG)
r = 3
for p, fill in (('POS', BLUE), ('SFA', GREEN)):
    section(ws, "A%d" % r, "ПРОДУКТ %s" % p); r += 1
    header(ws, r, ['Месяц','MRR на начало','Новые продажи','Расширения','Сокращения','Отток',
                   'MRR на конец','GRR','NRR','MoM']); r += 1
    first = r
    for a in series[p]:
        datarow(ws, r, [RUS[a['m']]+' 2026', a['start'], a['new'], a['exp'], a['con'], a['chu'],
                        f"=B{r}+C{r}+D{r}+E{r}+F{r}", f"=(B{r}+F{r}+E{r})/B{r}",
                        f"=(B{r}+F{r}+E{r}+D{r})/B{r}", f"=G{r}/B{r}-1"],
                fill=fill, fmts=[None,NUM,NUM,NUM,NUM,NUM,NUM,PCT,PCT,PCT]); r += 1
    totalrow(ws, r, ['ИТОГО ' + PERIOD, f"=B{first}", f"=SUM(C{first}:C{r-1})", f"=SUM(D{first}:D{r-1})",
                     f"=SUM(E{first}:E{r-1})", f"=SUM(F{first}:F{r-1})", f"=G{r-1}",
                     f"=AVERAGE(H{first}:H{r-1})", f"=AVERAGE(I{first}:I{r-1})", f"=G{r-1}/B{first}-1"],
             fmts=[None,NUM,NUM,NUM,NUM,NUM,NUM,PCT,PCT,PCT]); r += 2

section(ws, "A%d" % r, "РАЗОВАЯ ВЫРУЧКА ПОМЕСЯЧНО"); r += 1
header(ws, r, ['Месяц','POS, сум','SFA, сум','ВСЕГО','Доля в выручке POS','Доля в выручке SFA']); r += 1
first = r
for i, m in enumerate(MON):
    datarow(ws, r, [RUS[m]+' 2026', ONCE_M['POS'][i], ONCE_M['SFA'][i], f"=B{r}+C{r}",
                    ONCE_M['POS'][i]/(ONCE_M['POS'][i]+series['POS'][i]['end']),
                    ONCE_M['SFA'][i]/(ONCE_M['SFA'][i]+series['SFA'][i]['end'])],
            fill=YELLOW, fmts=[None,NUM,NUM,NUM,PCT,PCT]); r += 1
totalrow(ws, r, ['ИТОГО ' + PERIOD, f"=SUM(B{first}:B{r-1})", f"=SUM(C{first}:C{r-1})", f"=B{r}+C{r}", None, None],
         fmts=[None,NUM,NUM,NUM,None,None]); r += 1
datarow(ws, r, ['В среднем в месяц', f"=B{r-1}/{NM}", f"=C{r-1}/{NM}", f"=B{r}+C{r}", None, None],
        fill=GREY, fmts=[None,NUM,NUM,NUM,None,None]); r += 2
notes(ws, r, [
 "GRR = (MRR на начало − отток − сокращения) ÷ MRR на начало. NRR — то же плюс расширения. Новые продажи в удержание не входят.",
 "Разовая выручка POS: %s млн в феврале → %s млн в последнем месяце периода (%s) — следствие снижения числа новых подключений POS. "
 "У SFA разовая выручка неравномерна и зависит от отдельных проектов доработки и производства: %s млн в последнем месяце."
 % (mln(ONCE_M['POS'][0]), mln(ONCE_M['POS'][-1]), LAST_RU_L, mln(ONCE_M['SFA'][-1])),
 "Доля в выручке считается к сумме MRR на конец месяца и разовой выручки этого месяца.",
])

wb._sheets = [wb["Ключевые метрики"], wb["Помесячно 2026"]]
OUT = OUTP("Ключевые_метрики_SaaS_%s_с_разовой.xlsx" % SUF)
wb.save(OUT)
print("saved", OUT)
res = {'MRR': MRR, 'ONCE_MO': ONCE_MO, 'REV_TOT': REV_TOT, 'SRV': SRV, 'LIC': LIC, 'ARPA': ARPA, 'ARPU': ARPU, 'RPL': RPL,
       'GM_SUB': GM_SUB, 'GM_TOT': GM_TOT, 'COGS': COGS, 'CH_CLOSED': CH_CLOSED, 'CHURN_SUB': CHURN_SUB, 'CH_CONTR': CH_CONTR, 'EXPN': EXPN,
       'LIFE': LIFE, 'LTV': LTV, 'LTV_SUB': LTV_SUB, 'LTV_COH': LTV_COH, 'CAC': CAC, 'CAC_IMPL': CAC_IMPL, 'CAC_PAID': CAC_PAID,
       'CAC_SALES': CAC_SALES, 'CAC_ONB': CAC_ONB, 'NEW': NEW, 'ONCE': ONCE, 'MRR0': MRR0,
       'LTVCAC': {p: LTV[p]/CAC[p] for p in PR}, 'PAYBACK': {p: (CAC[p]-GP_ONCE[p])/GP_CYCLE[p] for p in PR},
       'GRR': {p: g(aug[p]) for p in PR}, 'NRR': {p: n(aug[p]) for p in PR}, 'MOM': {p: MRR[p]/jul[p]['end']-1 for p in PR},
       'series': series, 'ONCE_M': ONCE_M}
json.dump(res, open(J('res_metrics'), 'w'), ensure_ascii=False)
for p in PR:
    print("%s: MRR %.0f | разовая/мес %.0f | ARPA %.0f | GM %.1f%%/%.1f%% | отток %.2f%% | срок %.1f | LTV %.0f | CAC %.0f (без внедр. %.0f, платный %.0f) | LTV:CAC %.2f | payback %.1f"
          % (p, MRR[p], ONCE_MO[p], ARPA[p], GM_SUB[p]*100, GM_TOT[p]*100, CH_CLOSED[p]*100, LIFE[p], LTV[p], CAC[p], CAC_IMPL[p], CAC_PAID[p], LTV[p]/CAC[p], (CAC[p]-GP_ONCE[p])/GP_CYCLE[p]))
