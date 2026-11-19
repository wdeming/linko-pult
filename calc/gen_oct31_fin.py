# -*- coding: utf-8 -*-
"""Финансы для банка на срез: 1) P&L и EBITDA по месяцам, 2) деньги, расход и запас с займами, 3) признанная / выставленная / полученная выручка,
отложенная выручка и дебиторка, 4) сегменты клиентов по числу лицензий и когорты по деньгам. Excel + json для пульта."""
import json, os, sqlite3, collections, datetime, calendar
import openpyxl
from openpyxl.utils import get_column_letter as L
from openpyxl.styles import Font
from style_lib import *
from cfg_cut import *
K = 12000; D = datetime.date.fromisoformat
F = json.load(open(J('fin'), encoding='utf-8')); CASH = json.load(open(J('cash'), encoding='utf-8')); EXP = json.load(open(J('exp'), encoding='utf-8'))
COST = json.load(open(J('cost'), encoding='utf-8'))
ME = lambda m: datetime.date(int(m[:4]), int(m[5:]), calendar.monthrange(int(m[:4]), int(m[5:]))[1]).isoformat()
def split(a, d1, d2):
    x, y = D(d1), D(d2)
    if y < x: x, y = y, x
    n = (y - x).days + 1; res = collections.defaultdict(float); cur = x
    while cur <= y:
        e = min(y, cur.replace(day=calendar.monthrange(cur.year, cur.month)[1])); res[cur.strftime('%Y-%m')] += a * ((e - cur).days + 1) / n; cur = e + datetime.timedelta(days=1)
    return res
MS = [m for m in MON if m >= '2026-04']            # P&L и деньги — с апреля (помесячный учёт расходов)
MR = MON                                            # выручка — с февраля
RU = lambda m: '%s %s' % (RUS[m][:3].lower() if len(RUS[m]) > 4 else RUS[m].lower(), m[2:4])
# ---------- выручка: выставлено и признано ----------
SFA_P, POS_P = {25, 26, 41, 42}, {24, 40, 43, 32}
kind = lambda pid: 'Подписки SFA' if pid in SFA_P else 'Подписки POS' if pid in POS_P else 'Разовые работы и прочее'
INV = collections.defaultdict(float); REC = collections.defaultdict(float); INVK = collections.defaultdict(float); DEF = collections.defaultdict(float)
for x in F['inv']:
    k = kind(x['pid']); dm = x['d'][:7]
    if x['d'] <= CUT: INV[dm] += x['a']; INVK[(k, dm)] += x['a']
    sp = split(x['a'], x['d1'], x['d2']) if x['d1'] and x['d2'] else {dm: x['a']}
    for m, v in sp.items(): REC[(k, m)] += v
    for m in MR:      # отложенная выручка на конец месяца: выставлено до конца месяца, а признаётся позже
        if dm <= m: DEF[m] += sum(v for mm, v in sp.items() if mm > m)
recm = lambda m: sum(REC[(k, m)] for k in ('Подписки SFA', 'Подписки POS', 'Разовые работы и прочее'))
# ---------- MRR на конец месяца по журналу (POS + SFA) ----------
SUBS = F['subs']; LOG = F['log']
cum_sub = collections.defaultdict(lambda: collections.defaultdict(float))   # подписка -> месяц -> движение
for r in LOG: cum_sub[r['o']][r['m']] += r['a']
def mrr_at(o, m):   # MRR подписки на конец месяца m
    return sum(v for mm, v in cum_sub[o].items() if mm <= m)
MRR_END = {m: sum(mrr_at(o, m) for o in cum_sub if SUBS.get(str(o), {}).get('prod') in ('POS', 'SFA')) for m in MR}
# ---------- поступления, комиссии ----------
IN = collections.defaultdict(float); COM = collections.defaultdict(float)
for x in CASH:
    if x['pt'] == 'inbound': IN[x['m']] += x['s']
    else: COM[x['m']] += -x['s']
# ---------- расходы: за период по видам и по дате оплаты ----------
EP = collections.defaultdict(float); EPAY = collections.defaultdict(float); CAPEX = collections.defaultdict(float)
CAPCAT = {'Техника и мебель'}; HOSTCAT = 'Хостинг и инфраструктура'   # хостинг уже в себестоимости
for x in EXP:
    if x['pay'] <= CUT: EPAY[x['pay'][:7]] += x['amt']
    for m, v in split(x['amt'], x['d1'], x['d2']).items():
        if x['cat'] in CAPCAT: CAPEX[m] += v
        elif x['cat'] != HOSTCAT: EP[(x['plan'], m)] += v
COGS_FOT = {m: sum(COST['FOT_M'][k].get(m, 0) for k in ('SUP', 'ONP', 'ONS', 'KA')) for m in MS}
HOST = {m: COST['HOST_M'].get(m, 0) for m in MS}
# ---------- займы из реестра долгов пульта ----------
dbp = os.environ.get('LINKO_PULT_DB') or os.path.abspath(os.path.join(os.getcwd(), '..', 'data', 'pult.db'))
LT = collections.defaultdict(float); LP = collections.defaultdict(float); DEBT_END = {}; LOANS_OK = os.path.exists(dbp)
if LOANS_OK:
    c = sqlite3.connect(dbp)
    ops = c.execute("""select o.op, o.date, o.amount, d.currency from debt_op o join debt d on d.id=o.debt_id where o.void=0 and d.void=0""").fetchall()
    for op, dt, a, cur in ops:
        v = a * (K if cur == 'USD' else 1)
        (LT if op == 'take' else LP)[dt[:7]] += v
    for m in MR: DEBT_END[m] = sum(a * (K if cur == 'USD' else 1) * (1 if op == 'take' else -1) for op, dt, a, cur in ops if dt <= ME(m))
# ---------- остатки денег по месяцам ----------
ACC = collections.defaultdict(lambda: {'name': '', 'cur': 'UZS', 'mv': collections.defaultdict(float)})
for x in F['cash_acc']:
    a = ACC[x['code']]; a['name'] = x['name']; a['cur'] = x['cur']; a['mv'][x['m']] += (x['amc'] * K if x['cur'] == 'USD' else x['bal'])
bal_at = lambda code, m: sum(v for mm, v in ACC[code]['mv'].items() if mm <= m)
CASH_END = {m: sum(bal_at(code, m) for code in ACC) for m in MON + ['2026-03']}
prev = lambda m: MON[MON.index(m) - 1] if MON.index(m) > 0 else '2026-01'
# ---------- сегменты по лицензиям ----------
CL = collections.defaultdict(lambda: {'lic': 0.0, 'kas': 0.0, 'sfa': False, 'subs': []})
for o, s in SUBS.items():
    if s['prod'] not in ('POS', 'SFA') or not s['p']: continue
    c2 = CL[s['p']]; c2['subs'].append(int(o)); c2['lic'] += s['lic'] if s['prod'] == 'SFA' else 0; c2['kas'] += s['kas'] if s['prod'] == 'POS' else 0
    c2['sfa'] = c2['sfa'] or s['prod'] == 'SFA'
def seg(c2):
    if c2['sfa']:
        l = c2['lic']
        return 'SFA: лицензии не указаны' if l <= 0 else 'SFA: 1–5 агентов' if l <= 5 else 'SFA: 6–20 агентов' if l <= 20 else 'SFA: 21–50 агентов' if l <= 50 else 'SFA: 51+ агентов'
    return 'POS: 1 касса' if c2['kas'] <= 1 else 'POS: 2+ кассы'
SEGS = ['SFA: 1–5 агентов', 'SFA: 6–20 агентов', 'SFA: 21–50 агентов', 'SFA: 51+ агентов', 'SFA: лицензии не указаны', 'POS: 1 касса', 'POS: 2+ кассы']
M0 = '2026-02'; NMON = MON.index(LAST_M) - MON.index(M0)
L3 = MON[max(0, MON.index(LAST_M) - 3)]   # MRR 3 месяца назад
SG = {s: collections.defaultdict(float) for s in SEGS}
for p, c2 in CL.items():
    s = seg(c2); g = SG[s]
    m0 = sum(mrr_at(o, M0) for o in c2['subs']); mc = sum(mrr_at(o, LAST_M) for o in c2['subs']); m3 = sum(mrr_at(o, L3) for o in c2['subs'])
    if mc > 0: g['n'] += 1; g['mrr'] += mc; g['lic'] += c2['lic'] if c2['sfa'] else c2['kas']
    if m0 > 0:
        g['n0'] += 1; g['mrr0'] += m0; g['mrr0_now'] += max(mc, 0); g['grr_num'] += min(max(mc, 0), m0); g['lost_n'] += 1 if mc <= 0 else 0
    if m3 > 0: g['mrr3'] += m3
    for o in c2['subs']:
        for mm, v in cum_sub[o].items():
            if L3 < mm <= LAST_M and v < 0: g['loss3'] += -v
SEGS = [s for s in SEGS if SG[s]['n'] or SG[s]['n0']]
TOT = collections.defaultdict(float)
for s in SEGS:
    for k2, v in SG[s].items(): TOT[k2] += v
# ---------- когорты по деньгам (подписки по месяцу создания) ----------
first = {}
for r in LOG:
    if r['et'] == '0_creation' and SUBS.get(str(r['o']), {}).get('prod') in ('POS', 'SFA'): first[r['o']] = min(first.get(r['o'], '9999'), r['m'])
COH = {}
for m in [x for x in MON if x >= M0]:
    ss = [o for o, fm in first.items() if fm == m]
    st = sum(mrr_at(o, m) for o in ss)
    if st > 0: COH[m] = {'n': len(ss), 'start': st, 'k': [sum(mrr_at(o, mm) for o in ss) / st for mm in MON if mm >= m]}
# ---------- дебиторка ----------
BK = ['не просрочено', 'до 30 дней', '31–90 дней', '91–180 дней', 'больше 180 дней']; AR = collections.defaultdict(float); ARN = collections.defaultdict(int)
for x in F['ar']:
    dd = (D(CUT) - D(x['due'])).days; b = BK[0] if dd <= 0 else BK[1] if dd <= 30 else BK[2] if dd <= 90 else BK[3] if dd <= 180 else BK[4]
    AR[b] += x['res']; ARN[b] += 1
# ================= Excel =================
wb = openpyxl.Workbook(); wb.remove(wb.active)
fm = lambda n: [None] + [NUM] * n
# ---- 1. P&L ----
ws = wb.create_sheet('P&L и EBITDA'); n = len(MS); LC = L(1 + n)
title(ws, 'A1', 'P&L и EBITDA по месяцам, апрель–%s 2026 (на %s), сум' % (LAST_RU_L, CUT_RU))
header(ws, 3, ['Статья'] + ['%s 2026' % RUS[m] for m in MS] + ['ИТОГО'])
r = 4; ROW = {}
def line(lab, vals, key=None, total=True, fmt=NUM, bold=False, fill=None):
    global r
    v = [lab] + vals + (['=SUM(B%d:%s%d)' % (r, LC, r)] if total else [None])
    (totalrow if bold else datarow)(ws, r, v, fmts=[None] + [fmt] * (n + 1)) if bold else datarow(ws, r, v, fmts=[None] + [fmt] * (n + 1), fill=fill)
    if key: ROW[key] = r
    r += 1
def frm(expr): return ['=' + expr.replace('#', L(cc)) for cc in range(2, 2 + n)]
section(ws, 'A%d' % r, 'ВЫРУЧКА (признанная — по периоду услуги)'); r += 1
for k in ('Подписки SFA', 'Подписки POS', 'Разовые работы и прочее'): line('    ' + k, [round(REC[(k, m)]) for m in MS], k)
line('ИТОГО ВЫРУЧКА', frm('#%d+#%d+#%d' % (ROW['Подписки SFA'], ROW['Подписки POS'], ROW['Разовые работы и прочее'])), 'rev', bold=True)
section(ws, 'A%d' % r, 'СЕБЕСТОИМОСТЬ'); r += 1
line('    ФОТ поддержки, внедрения и ключевых клиентов', [round(COGS_FOT[m]) for m in MS], 'cfot')
line('    Хостинг и инфраструктура', [round(HOST[m]) for m in MS], 'chost')
line('ИТОГО СЕБЕСТОИМОСТЬ', frm('#%d+#%d' % (ROW['cfot'], ROW['chost'])), 'cogs', bold=True)
line('ВАЛОВАЯ ПРИБЫЛЬ', frm('#%d-#%d' % (ROW['rev'], ROW['cogs'])), 'gp', bold=True)
line('    Валовая маржа, %', ['=IF(#%d=0,0,#%d/#%d)'.replace('#', L(cc)) % (ROW['rev'], ROW['gp'], ROW['rev']) for cc in range(2, 2 + n)], 'gm', total=False, fmt=PCT)
section(ws, 'A%d' % r, 'ОПЕРАЦИОННЫЕ РАСХОДЫ (за период, без того, что в себестоимости)'); r += 1
OPX = [('Персонал (остальные отделы, налоги ЗП, обед, аутсорс)', {m: EP[('Затраты на персонал', m)] - COGS_FOT[m] for m in MS}),
       ('Административные расходы', {m: EP[('Административные расходы', m)] for m in MS}),
       ('IT услуги (без хостинга)', {m: EP[('IT услуги', m)] for m in MS}),
       ('Операционные расходы', {m: EP[('Операционные расходы', m)] for m in MS}),
       ('Услуги связи и интернет', {m: EP[('Услуги связи и интернет', m)] for m in MS}),
       ('Комиссии платёжных систем (Payme, Click, Paynet)', {m: COM[m] for m in MS})]
r0 = r
for lab, v in OPX: line('    ' + lab, [round(v[m]) for m in MS])
line('ИТОГО ОПЕРАЦИОННЫЕ РАСХОДЫ', ['=SUM(%s%d:%s%d)' % (L(cc), r0, L(cc), r - 1) for cc in range(2, 2 + n)], 'opex', bold=True)
line('EBITDA', frm('#%d-#%d' % (ROW['gp'], ROW['opex'])), 'ebitda', bold=True)
line('    EBITDA, % от выручки', ['=IF(#%d=0,0,#%d/#%d)'.replace('#', L(cc)) % (ROW['rev'], ROW['ebitda'], ROW['rev']) for cc in range(2, 2 + n)], total=False, fmt=PCT)
line('Капитальные расходы (техника и мебель) — не входят в EBITDA', [round(CAPEX[m]) for m in MS], 'capex')
r += 1; section(ws, 'A%d' % r, 'Справочно'); r += 1
line('    Выставлено счетов', [round(INV[m]) for m in MS])
line('    MRR на конец месяца', [round(MRR_END[m]) for m in MS], total=False)
notes(ws, r + 1, [
  'Выручка — признанная: каждая строка счета клиенту разнесена по дням периода услуги (поля периода в строке счета Odoo). Годовая предоплата признаётся помесячно, а не в месяце счета. Строки без периода — в месяце счета. Учитываются и счета, выставленные после %s за услуги до этой даты.' % CUT_RU,
  'Себестоимость — факт из реестра «Расходы» Odoo: ФОТ отделов поддержки, внедрения и ключевых клиентов (оклад, KPI, прочие выплаты за период + доля налогов ЗП) и хостинг за период. Файл «Себестоимость_поддержки_внедрения_инфраструктуры_%s.xlsx».' % SUF,
  'Операционные расходы — реестр «Расходы» за период (сумма делится по дням периода) минус то, что уже в себестоимости; комиссии платёжных систем — исходящие платежи Payme / Click / Paynet.',
  'EBITDA = валовая прибыль − операционные расходы. Покупка техники и мебели — капитальные расходы, показаны отдельно под EBITDA. Процентов по займам в реестре расходов нет.',
  'Суммы в долларах — по управленческому курсу 12 000 сум/$ (методологическое решение менеджмента). Данные с апреля 2026: до этого расходы в Odoo по месяцам не разнесены.'])
widths(ws, {'A': 56, **{L(cc): 15 for cc in range(2, 3 + n)}}); ws.freeze_panes = 'B4'
PL = {m: {'rev': recm(m), 'cogs': COGS_FOT[m] + HOST[m], 'opex': sum(v[m] for _, v in OPX)} for m in MS}
for m in MS: PL[m]['gp'] = PL[m]['rev'] - PL[m]['cogs']; PL[m]['ebitda'] = PL[m]['gp'] - PL[m]['opex']
# ---- 2. Деньги ----
ws = wb.create_sheet('Деньги и запас')
title(ws, 'A1', 'Деньги, расход и запас по месяцам (на %s), сум' % CUT_RU)
header(ws, 3, ['Статья'] + ['%s 2026' % RUS[m] for m in MS] + ['ИТОГО']); r = 4; ROW = {}
line('Остаток денег на начало месяца', [round(CASH_END[prev(m)]) for m in MS], 'o', total=False)
line('    + Поступления от клиентов', [round(IN[m]) for m in MS], 'in')
line('    − Комиссии платёжных систем', [round(-COM[m]) for m in MS], 'com')
line('    − Расходы по дате оплаты', [round(-EPAY[m]) for m in MS], 'exp')
line('ОПЕРАЦИОННЫЙ ДЕНЕЖНЫЙ ПОТОК', frm('#%d+#%d+#%d' % (ROW['in'], ROW['com'], ROW['exp'])), 'op', bold=True)
line('    + Займы получены (реестр долгов пульта)', [round(LT[m]) for m in MS], 'lt')
line('    − Займы возвращены', [round(-LP[m]) for m in MS], 'lp')
line('    Прочие движения и расхождения учёта', ['=#%d-#%d-#%d-#%d-#%d'.replace('#', L(cc)) % (r + 1, ROW['o'], ROW['op'], ROW['lt'], ROW['lp']) for cc in range(2, 2 + n)], 'oth')
line('Остаток денег на конец месяца', [round(CASH_END[m]) for m in MS], 'e', total=False, bold=True)
r += 1
line('Средний операционный поток за 3 месяца', ['=AVERAGE(%s%d:%s%d)' % (L(max(2, cc - 2)), ROW['op'], L(cc), ROW['op']) for cc in range(2, 2 + n)], 'avg', total=False)
datarow(ws, r, ['Запас денег, месяцев (при отрицательном потоке)'] + ['=IF(%s%d>=0,"поток положительный",%s%d/-%s%d)' % (L(cc), ROW['avg'], L(cc), ROW['e'], L(cc), ROW['avg']) for cc in range(2, 2 + n)], fmts=[None] + ['0.0'] * n, blue_inputs=False); r += 1
line('Долг по займам на конец месяца (реестр долгов пульта)', [round(DEBT_END.get(m, 0)) for m in MS], total=False)
r += 1; header(ws, r, ['Остаток на %s по счетам' % CUT_RU, 'Сумма, сум', 'Валюта']); r += 1; ra = r
for code in sorted(ACC, key=lambda c3: -bal_at(c3, LAST_M)):
    datarow(ws, r, ['%s · %s' % (code, ACC[code]['name']), round(bal_at(code, LAST_M)), ACC[code]['cur']], fmts=[None, NUM, None], blue_inputs=False); r += 1
totalrow(ws, r, ['ИТОГО денег', '=SUM(B%d:B%d)' % (ra, r - 1)], fmts=[None, NUM])
notes(ws, r + 2, [
  'Деньги — остатки денежных счетов Odoo (банки, Click / Payme / Paynet на счёте банка, карты, наличные). Не включены «Касса Азиз» (учёт долга перед учредителями, не деньги компании) и FOTON.',
  'Валютные счета — по сумме в долларах × 12 000. По счёту 101505 «Наличные валюта» в сумах в Odoo числится больше, чем в долларах: часть проводок без валюты — требует сверки с картой.',
  'Займы — из реестра долгов на пульте (ведётся вручную). «Прочие движения» — разница между изменением остатка и потоками выше: расходы, оплаченные учредителями напрямую (через «Кассу Азиз»), переводы, неразнесённые проводки.',
  'Запас = остаток денег на конец месяца ÷ средний отрицательный операционный поток за 3 месяца.'])
widths(ws, {'A': 54, **{L(cc): 15 for cc in range(2, 3 + n)}}); ws.freeze_panes = 'B4'
CS = {m: {'open': CASH_END[prev(m)], 'in': IN[m], 'com': COM[m], 'exp': EPAY[m], 'op': IN[m] - COM[m] - EPAY[m], 'lt': LT[m], 'lp': LP[m], 'end': CASH_END[m], 'debt': DEBT_END.get(m)} for m in MS}
for m in MS: CS[m]['oth'] = CS[m]['end'] - CS[m]['open'] - CS[m]['op'] - CS[m]['lt'] + CS[m]['lp']
last3 = [CS[m]['op'] for m in MS[-3:]]; avg3 = sum(last3) / len(last3)
# ---- 3. Выручка ----
ws = wb.create_sheet('Выручка и отложенная')
title(ws, 'A1', 'MRR, признанная, выставленная и полученная выручка по месяцам (на %s), сум' % CUT_RU)
header(ws, 3, ['Месяц', 'MRR на конец месяца', 'Признано (выручка периода)', 'Выставлено счетов', 'Получено денег', 'Отложенная выручка на конец месяца', 'в т.ч. признано SFA', 'в т.ч. признано POS', 'Выставлено SFA', 'Выставлено POS'])
r = 4
for m in MR:
    datarow(ws, r, ['%s 2026' % RUS[m], round(MRR_END[m]), round(recm(m)), round(INV[m]), round(IN[m]), round(DEF.get(m, 0)), round(REC[('Подписки SFA', m)]), round(REC[('Подписки POS', m)]),
                    round(INVK[('Подписки SFA', m)]), round(INVK[('Подписки POS', m)])], fmts=fm(9)); r += 1
totalrow(ws, r, ['ИТОГО'] + [None] + ['=SUM(%s4:%s%d)' % (L(cc), L(cc), r - 1) for cc in (3, 4, 5)] + [None] + ['=SUM(%s4:%s%d)' % (L(cc), L(cc), r - 1) for cc in (7, 8, 9, 10)], fmts=fm(9))
r += 2; header(ws, r, ['Дебиторка на %s (остаток на сегодня)' % CUT_RU, 'Счетов', 'Сумма, сум']); r += 1; r1 = r
for b in BK: datarow(ws, r, [b, ARN[b], round(AR[b])], fmts=[None, NUM, NUM], blue_inputs=False); r += 1
totalrow(ws, r, ['ИТОГО неоплаченных счетов', '=SUM(B%d:B%d)' % (r1, r - 1), '=SUM(C%d:C%d)' % (r1, r - 1)], fmts=[None, NUM, NUM]); r += 1
datarow(ws, r, ['Поступления, не привязанные к счетам (на ту же дату)', F['unmatched']['n'], round(F['unmatched']['s'])], fmts=[None, NUM, NUM], blue_inputs=False)
notes(ws, r + 2, [
  'Признано — строки счетов разнесены по дням периода услуги; выставлено — по дате счета; получено — проведённые платежи клиентов по дате (без внутренних переводов).',
  'Скачки «выставлено» по SFA — годовые и полугодовые предоплаты: счёт выставляется разом, а выручка признаётся по месяцам. Отложенная выручка = всё выставленное минус всё признанное на конец месяца (обязательство перед клиентами).',
  'Автоматическое признание отложенной выручки в Odoo остановилось в апреле 2026 (остаток на счёте 201000 не списывается) — здесь признание посчитано заново по периодам в строках счетов, от проводок Odoo не зависит.',
  'Дебиторка завышена: в Odoo много поступлений не привязано к счетам (строка выше), поэтому оплаченные счета числятся неоплаченными. Реальную дебиторку покажет сверка платежей со счетами в Odoo.'])
widths(ws, {'A': 46, **{L(cc): 17 for cc in range(2, 11)}}); ws.freeze_panes = 'B4'
# ---- 4. Сегменты ----
ws = wb.create_sheet('Сегменты клиентов')
title(ws, 'A1', 'Сегменты клиентов по размеру: удержание и отток по деньгам (на %s)' % CUT_RU)
header(ws, 3, ['Сегмент', 'Клиентов на срез', 'MRR на срез', 'Доля MRR', 'Лицензий / касс', 'Клиентов на 28.02', 'MRR на 28.02', 'MRR этих клиентов на срез', 'NRR с 28.02', 'GRR с 28.02', 'NRR в пересчёте на год', 'Ушли клиентов с 28.02', 'Потери MRR за 3 мес. (отток + сокращения)'])
r = 4; SR = {}
def segrow(lab, g, bold=False):
    global r
    nrr = g['mrr0_now'] / g['mrr0'] if g['mrr0'] else None; grr = g['grr_num'] / g['mrr0'] if g['mrr0'] else None
    v = [lab, g['n'], round(g['mrr']), g['mrr'] / TOT['mrr'] if TOT['mrr'] else 0, g['lic'], g['n0'], round(g['mrr0']), round(g['mrr0_now']), nrr, grr,
         nrr ** (12 / NMON) if nrr and NMON else None, (g['lost_n'] / g['n0']) if g['n0'] else None, (g['loss3'] / g['mrr3']) if g['mrr3'] else None]
    (totalrow if bold else datarow)(ws, r, v, fmts=[None, NUM, NUM, PCT, NUM, NUM, NUM, NUM, PCT, PCT, PCT, PCT, PCT]); r += 1
    return {'n': g['n'], 'mrr': g['mrr'], 'nrr': nrr, 'grr': grr, 'nrr_y': v[10], 'lost': v[11], 'loss3': v[12], 'lic': g['lic']}
for s in SEGS: SR[s] = segrow(s, SG[s])
SR['Всего'] = segrow('ИТОГО', TOT, True)
r += 1; header(ws, r, ['Когорта (месяц подключения подписки)', 'Подписок', 'MRR на конец 1-го месяца'] + ['М%d' % i for i in range(NMON + 1)]); r += 1
for m, cdata in COH.items():
    datarow(ws, r, ['%s 2026' % RUS[m], cdata['n'], round(cdata['start'])] + cdata['k'], fmts=[None, NUM, NUM] + [PCT] * len(cdata['k']), blue_inputs=False); r += 1
notes(ws, r + 1, [
  'Размер клиента — число лицензий SFA (агентов) во всех его подписках SFA; клиенты только с POS — по числу касс. Число лицензий — по текущему составу подписок (у ушедших — по последнему).',
  'NRR / GRR — по клиентам, у которых был MRR на 28.02.2026: их MRR на срез ÷ MRR на 28.02 (GRR — без учёта роста сверх 28.02). Это %d месяцев; «в пересчёте на год» = NRR^(12/%d). Честный годовой NRR будет с февраля 2027 — журнал MRR в Odoo ведётся с конца 2025.' % (NMON, NMON),
  'Потери за 3 мес. — все отрицательные движения MRR (отток и сокращения) за последние 3 месяца ÷ MRR сегмента 3 месяца назад.',
  'Когорты — подписки по месяцу создания; М0 = 100%% MRR на конец месяца подключения, дальше — доля этого MRR на конец каждого следующего месяца (с расширениями и сокращениями). Всё по журналу MRR Odoo.'])
widths(ws, {'A': 34, **{L(cc): 15 for cc in range(2, 16)}}); ws.freeze_panes = 'B4'
fn = OUTP('Финансы_для_банка_%s_2026_%s.xlsx' % (PERIOD.replace('–', '-'), SUF)); wb.save(fn)
res = {'CAPEX': {m: CAPEX[m] for m in MS}, 'MS': MS, 'MR': MR, 'PL': PL, 'REC': {'%s|%s' % kk: v for kk, v in REC.items() if kk[1] in MR}, 'INV': {m: INV[m] for m in MR}, 'IN': {m: IN[m] for m in MR}, 'MRR': MRR_END,
       'DEF': {m: DEF.get(m, 0) for m in MR}, 'OPX': {lab: v for lab, v in OPX}, 'COGS_FOT': COGS_FOT, 'HOST': HOST, 'CS': CS, 'AVG3': avg3, 'LOANS_OK': LOANS_OK,
       'ACC': {('%s · %s' % (code, ACC[code]['name'])): bal_at(code, LAST_M) for code in ACC}, 'AR': {b: AR[b] for b in BK}, 'ARN': {b: ARN[b] for b in BK}, 'UNM': F['unmatched'],
       'SEG': SR, 'SEGS': SEGS, 'COH': COH, 'NMON': NMON}
json.dump(res, open(J('res_fin'), 'w'), ensure_ascii=False)
print(fn)
for m in MS: print(m, {k: round(v / 1e6, 1) for k, v in PL[m].items()}, '| деньги', {k: round((v or 0) / 1e6, 1) for k, v in CS[m].items()})
print('avg3', round(avg3 / 1e6, 1), 'loans', LOANS_OK)
for m in MR: print(m, 'MRR', round(MRR_END[m] / 1e6, 1), 'rec', round(recm(m) / 1e6, 1), 'inv', round(INV[m] / 1e6, 1), 'in', round(IN[m] / 1e6, 1), 'def', round(DEF.get(m, 0) / 1e6, 1))
for s, v in SR.items(): print(s, {k: (round(x, 3) if isinstance(x, float) and abs(x) < 10 else round(x / 1e6, 1) if isinstance(x, float) else x) for k, x in v.items()})
print('AR', {b: round(AR[b] / 1e6, 1) for b in BK}, F['unmatched'])
for m, cdata in COH.items(): print(m, cdata['n'], round(cdata['start'] / 1e6, 1), [round(x, 2) for x in cdata['k']])
