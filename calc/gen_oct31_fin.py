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
import fin_groups as FG
dbp = os.environ.get('LINKO_PULT_DB') or os.path.abspath(os.path.join(os.getcwd(), '..', 'data', 'pult.db'))
_hp = [os.environ.get('LINKO_FIN_HIST') or '', os.path.join(os.path.dirname(dbp), 'fin_hist.json'), os.path.join(os.getcwd(), 'fin_hist.json')]
HP = next((x for x in _hp if x and os.path.exists(x)), None)
HIST = json.load(open(HP, encoding='utf-8')) if HP else None      # 2025 — файл финансиста, янв–мар 2026 — пакет (build_fin_hist.py)
HC = HIST['months'] if HIST else {}
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
LATE_IN = collections.defaultdict(float); LATE_OUT = collections.defaultdict(float)   # для мостика «по периодам услуги» → «как в Odoo»
for x in F['inv']:
    k = kind(x['pid']); dm = x['d'][:7]
    if x['d'] <= CUT: INV[dm] += x['a']; INVK[(k, dm)] += x['a']
    sp = split(x['a'], x['d1'], x['d2']) if x['d1'] and x['d2'] else {dm: x['a']}
    for m, v in sp.items():
        REC[(k, m)] += v
        if m < dm:
            LATE_OUT[m] += v                                # услуга месяца m, счёт выставлен позже
            if x['d'] <= CUT: LATE_IN[dm] += v              # счёт месяца dm за прошлый период
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
EG = collections.defaultdict(float)     # единые группы расходов (fin_groups) — для истории с 2025
for x in EXP:
    if x['pay'] <= CUT: EPAY[x['pay'][:7]] += x['amt']
    for m, v in split(x['amt'], x['d1'], x['d2']).items():
        EG[(FG.group(x['cat']), m)] += v
        if x['cat'] in CAPCAT: CAPEX[m] += v
        elif x['cat'] != HOSTCAT: EP[(x['plan'], m)] += v
COGS_FOT = {m: sum(COST['FOT_M'][k].get(m, 0) for k in ('SUP', 'ONP', 'ONS', 'KA')) for m in MS}
HOST = {m: COST['HOST_M'].get(m, 0) for m in MS}
# ---------- займы из реестра долгов пульта ----------
LT = collections.defaultdict(float); LP = collections.defaultdict(float); DEBT_END = {}; LOANS_OK = os.path.exists(dbp)
LOANK, PAYK = ('Займ', 'Учредитель'), 'Зарплата / сотрудники'; DB_BAL = collections.defaultdict(float)
if LOANS_OK:     # в движении денег — только займы; долги поставщикам и по ФОТ оплачиваются через реестр расходов и попадают в баланс
    c = sqlite3.connect('file:%s?mode=ro' % dbp, uri=True)
    ops = c.execute("""select o.op, o.date, o.amount, d.currency, d.kind from debt_op o join debt d on d.id=o.debt_id where o.void=0 and d.void=0""").fetchall()
    loans = [(op, dt, a, cur) for op, dt, a, cur, kd in ops if kd in LOANK]
    for op, dt, a, cur in loans:
        v = a * (K if cur == 'USD' else 1)
        (LT if op == 'take' else LP)[dt[:7]] += v
    for m in ['2026-01'] + MR: DEBT_END[m] = sum(a * (K if cur == 'USD' else 1) * (1 if op == 'take' else -1) for op, dt, a, cur in loans if dt <= ME(m))
    for op, dt, a, cur, kd in ops:
        if dt <= CUT: DB_BAL['loan' if kd in LOANK else 'pay' if kd == PAYK else 'sup'] += a * (K if cur == 'USD' else 1) * (1 if op == 'take' else -1)
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
M0 = '2026-01'; NMON = int(LAST_M[5:7]) - 1   # база подписок загружена в Odoo в январе 2026: 31.01 — самая ранняя полная точка
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
# ---------- NRR / GRR за весь доступный период: с 31.01.2026 (база подписок загружена в журнал MRR в январе) ----------
MLIST = ['2026-01'] + [m for m in MON if m <= LAST_M]; KY = 12 / NMON if NMON else None
MRR_END['2026-01'] = sum(mrr_at(o, '2026-01') for o in cum_sub if SUBS.get(str(o), {}).get('prod') in ('POS', 'SFA'))
RET = {}
for lab, pf in (('Всего', None), ('SFA', 'SFA'), ('POS', 'POS')):
    cp = collections.defaultdict(list)
    for o, s2 in SUBS.items():
        if s2['prod'] in ('POS', 'SFA') and s2['p'] and (pf is None or s2['prod'] == pf): cp[s2['p']].append(int(o))
    base0 = {p: v for p, v in ((p, sum(mrr_at(o, M0) for o in ss)) for p, ss in cp.items()) if v > 0}; s0 = sum(base0.values())
    curve = []
    for m in MLIST:
        now = {p: sum(mrr_at(o, m) for o in cp[p]) for p in base0}
        curve.append({'m': m, 'nrr': sum(max(v, 0) for v in now.values()) / s0, 'grr': sum(min(max(now[p], 0), base0[p]) for p in base0) / s0, 'lost': sum(1 for v in now.values() if v <= 0)})
    lc = curve[-1]
    RET[lab] = {'n0': len(base0), 'mrr0': s0, 'mrr_now': lc['nrr'] * s0, 'nrr': lc['nrr'], 'grr': lc['grr'], 'nrr_y': lc['nrr'] ** KY if KY else None, 'grr_y': lc['grr'] ** KY if KY else None,
                'lost': lc['lost'] / len(base0) if base0 else None, 'curve': curve}
# сокращения: «колебания» (позже вернули расширением той же подписки) и настоящие потери
SPL = {m: collections.defaultdict(float) for m in MLIST[1:]}
EV = collections.defaultdict(list); ORD = {'15_contraction': 0, '1_expansion': 1, '2_churn': 2, '0_creation': 3}
for r in LOG:
    if MLIST[1] <= r['m'] <= LAST_M and SUBS.get(str(r['o']), {}).get('prod') in ('POS', 'SFA'): EV[r['o']].append((r['m'], ORD.get(r['et'], 9), r['et'], r['a']))
for o, evs in EV.items():
    lots = []
    for m, _, et, a in sorted(evs):
        g = SPL[m]
        if et == '15_contraction': g['con'] += -a; lots.append([m, -a])
        elif et == '1_expansion':
            g['exp'] += a; left = a
            while lots and left > 0:
                tk = min(lots[0][1], left); SPL[lots[0][0]]['con_back'] += tk; g['exp_back'] += tk; left -= tk; lots[0][1] -= tk
                if lots[0][1] <= 1e-6: lots.pop(0)
        elif et == '2_churn': g['churn'] += -a; lots = []
        elif et == '0_creation': g['new'] += a
MONTHLY = {}
for i, m in enumerate(MLIST[1:]):
    g = SPL[m]; st = MRR_END[MLIST[i]]; cr = g['con'] - g['con_back']
    MONTHLY[m] = {'start': st, 'new': g['new'], 'exp': g['exp'], 'exp_back': g['exp_back'], 'con': g['con'], 'con_back': g['con_back'], 'con_real': cr, 'churn': g['churn'],
                  'grr': 1 - (g['churn'] + g['con']) / st if st else None, 'grr_adj': 1 - (g['churn'] + cr) / st if st else None, 'nrr': 1 + (g['exp'] - g['con'] - g['churn']) / st if st else None}
import math
_adj = [x['grr_adj'] for x in MONTHLY.values() if x['grr_adj']]; _rep = [x['grr'] for x in MONTHLY.values() if x['grr']]
GRR_ADJ_Y = math.prod(_adj) ** (12 / len(_adj)) if _adj else None; GRR_REP_Y = math.prod(_rep) ** (12 / len(_rep)) if _rep else None
CON_T = sum(x['con'] for x in MONTHLY.values()); CON_B = sum(x['con_back'] for x in MONTHLY.values())
# ---------- мостик: выручка по периодам услуги → как в Odoo (по дате счёта, вся сумма сразу) ----------
pm = lambda m: MON[MON.index(m) - 1] if MON.index(m) > 0 else '2026-01'
BR = {m: {'ddef': DEF.get(m, 0) - DEF.get(pm(m), 0), 'late_in': LATE_IN[m], 'late_out': LATE_OUT[m]} for m in MS}
for m in MS: assert abs(recm(m) + BR[m]['ddef'] + BR[m]['late_in'] - BR[m]['late_out'] - INV[m]) < 5, ('мостик не сходится', m)
for m in MS: EG[(FG.G_BANK, m)] += COM[m]
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
r += 1; section(ws, 'A%d' % r, 'ВЫРУЧКА КАК В ODOO — по дате счёта, вся сумма сразу (учётная политика компании)'); r += 1
line('Выручка по периодам услуги (строка ИТОГО ВЫРУЧКА)', frm('#%d' % ROW['rev']), 'b0')
line('    + рост оплаченного клиентами вперёд за будущие месяцы', [round(BR[m]['ddef']) for m in MS], 'b1')
line('    + счета месяца за услуги прошлых месяцев', [round(BR[m]['late_in']) for m in MS], 'b2')
line('    − услуги месяца, счёт на которые выставлен позже', [round(-BR[m]['late_out']) for m in MS], 'b3')
line('ВЫРУЧКА КАК В ODOO', frm('#%d+#%d+#%d+#%d' % (ROW['b0'], ROW['b1'], ROW['b2'], ROW['b3'])), 'rodoo', bold=True)
line('    проверка: выставлено счетов за месяц (Odoo)', [round(INV[m]) for m in MS])
line('EBITDA, если выручку брать как в Odoo', frm('#%d-#%d-#%d' % (ROW['rodoo'], ROW['cogs'], ROW['opex'])), 'eodoo')
r += 1; section(ws, 'A%d' % r, 'Справочно'); r += 1
line('    MRR на конец месяца', [round(MRR_END[m]) for m in MS], total=False)
notes(ws, r + 1, [
  'Выручка — признанная: каждая строка счета клиенту разнесена по дням периода услуги (поля периода в строке счета Odoo). Годовая предоплата признаётся помесячно, а не в месяце счета. Строки без периода — в месяце счета. Учитываются и счета, выставленные после %s за услуги до этой даты.' % CUT_RU,
  'Себестоимость — факт из реестра «Расходы» Odoo: ФОТ отделов поддержки, внедрения и ключевых клиентов (оклад, KPI, прочие выплаты за период + доля налогов ЗП) и хостинг за период. Файл «Себестоимость_поддержки_внедрения_инфраструктуры_%s.xlsx».' % SUF,
  'Операционные расходы — реестр «Расходы» за период (сумма делится по дням периода) минус то, что уже в себестоимости; комиссии платёжных систем — исходящие платежи Payme / Click / Paynet.',
  'EBITDA = валовая прибыль − операционные расходы. Покупка техники и мебели — капитальные расходы, показаны отдельно под EBITDA. Процентов по займам в реестре расходов нет.',
  'Главная выручка — по периоду услуги (как требует МСФО 15 и как смотрит банк). В Odoo по учётной политике компании выручка признаётся сразу в момент продажи — эта цифра показана ниже («Выручка как в Odoo») с мостиком: разница = рост оплаченного клиентами вперёд за будущие месяцы ± счета, выставленные не в месяце услуги. Обе цифры из одних и тех же счетов Odoo и сходятся до суммы.',
  'Суммы в долларах — по управленческому курсу 12 000 сум/$ (методологическое решение менеджмента). Данные с апреля 2026: до этого расходы в Odoo по месяцам не разнесены.'])
widths(ws, {'A': 56, **{L(cc): 15 for cc in range(2, 3 + n)}}); ws.freeze_panes = 'B4'
PL = {m: {'rev': recm(m), 'cogs': COGS_FOT[m] + HOST[m], 'opex': sum(v[m] for _, v in OPX)} for m in MS}
for m in MS:
    PL[m]['gp'] = PL[m]['rev'] - PL[m]['cogs']; PL[m]['ebitda'] = PL[m]['gp'] - PL[m]['opex']
    PL[m]['rev_odoo'] = INV[m]; PL[m]['ebitda_odoo'] = INV[m] - PL[m]['cogs'] - PL[m]['opex']
    assert abs(sum(EG[(g, m)] for g in FG.GROUPS) - PL[m]['cogs'] - PL[m]['opex']) < 5, ('группы расходов не сходятся с P&L', m)
# ---- 2. Деньги ----
MC = [m for m in ('2026-01', '2026-02', '2026-03') if 'cash' in HC.get(m, {})] + MS     # янв–мар — пакет (ДДС), с апреля — Odoo
CS = {}
for i, m in enumerate(MC):
    if m in MS: CS[m] = {'open': CASH_END[prev(m)], 'in': IN[m], 'com': COM[m], 'exp': EPAY[m], 'end': CASH_END[m], 'src': 'Odoo'}
    else:
        hc = HC[m]['cash']
        CS[m] = {'open': CS[MC[i - 1]]['end'] if i else 0.0, 'in': hc['in'], 'com': 0.0, 'exp': hc['pay'], 'end': CASH_END['2026-03'] if m == '2026-03' else hc['end'], 'src': 'пакет'}
    x = CS[m]; x['op'] = x['in'] - x['com'] - x['exp']; x['lt'] = LT[m]; x['lp'] = LP[m]; x['debt'] = DEBT_END.get(m)
    x['oth'] = x['end'] - x['open'] - x['op'] - x['lt'] + x['lp']
n = len(MC); LC = L(1 + n)
ws = wb.create_sheet('Деньги и запас')
title(ws, 'A1', 'Деньги, расход и запас по месяцам (на %s), сум' % CUT_RU)
header(ws, 3, ['Статья'] + ['%s 2026' % RUS.get(m, 'Январь' if m == '2026-01' else m) for m in MC] + ['ИТОГО']); r = 4; ROW = {}
datarow(ws, r, ['Источник'] + [CS[m]['src'] for m in MC], blue_inputs=False); r += 1
line('Остаток денег на начало месяца', [round(CS[m]['open']) for m in MC], 'o', total=False)
line('    + Поступления от клиентов', [round(CS[m]['in']) for m in MC], 'in')
line('    − Комиссии платёжных систем', [round(-CS[m]['com']) for m in MC], 'com')
line('    − Расходы по дате оплаты', [round(-CS[m]['exp']) for m in MC], 'exp')
line('ОПЕРАЦИОННЫЙ ДЕНЕЖНЫЙ ПОТОК', frm('#%d+#%d+#%d' % (ROW['in'], ROW['com'], ROW['exp'])), 'op', bold=True)
line('    + Займы получены (раздел «Долги» пульта)', [round(LT[m]) for m in MC], 'lt')
line('    − Займы возвращены', [round(-LP[m]) for m in MC], 'lp')
line('    Прочие движения и расхождения учёта', ['=#%d-#%d-#%d-#%d-#%d'.replace('#', L(cc)) % (r + 1, ROW['o'], ROW['op'], ROW['lt'], ROW['lp']) for cc in range(2, 2 + n)], 'oth')
line('Остаток денег на конец месяца', [round(CS[m]['end']) for m in MC], 'e', total=False, bold=True)
r += 1
line('Средний операционный поток за 3 месяца', ['=AVERAGE(%s%d:%s%d)' % (L(max(2, cc - 2)), ROW['op'], L(cc), ROW['op']) for cc in range(2, 2 + n)], 'avg', total=False)
datarow(ws, r, ['Запас денег, месяцев (при отрицательном потоке)'] + ['=IF(%s%d>=0,"поток положительный",%s%d/-%s%d)' % (L(cc), ROW['avg'], L(cc), ROW['e'], L(cc), ROW['avg']) for cc in range(2, 2 + n)], fmts=[None] + ['0.0'] * n, blue_inputs=False); r += 1
line('Долг по займам на конец месяца (раздел «Долги» пульта)', [round(DEBT_END.get(m, 0)) for m in MC], total=False)
r += 1; header(ws, r, ['Остаток на %s по счетам' % CUT_RU, 'Сумма, сум', 'Валюта']); r += 1; ra = r
for code in sorted(ACC, key=lambda c3: -bal_at(c3, LAST_M)):
    datarow(ws, r, ['%s · %s' % (code, ACC[code]['name']), round(bal_at(code, LAST_M)), ACC[code]['cur']], fmts=[None, NUM, None], blue_inputs=False); r += 1
totalrow(ws, r, ['ИТОГО денег', '=SUM(B%d:B%d)' % (ra, r - 1)], fmts=[None, NUM])
notes(ws, r + 2, [
  'Деньги — остатки денежных счетов Odoo (банки, Click / Payme / Paynet на счёте банка, карты, наличные). Не включены «Касса Азиз» (учёт долга перед учредителями, не деньги компании) и FOTON.',
  'Валютные счета — по сумме в долларах × 12 000. По счёту 101505 «Наличные валюта» в сумах в Odoo числится больше, чем в долларах: часть проводок без валюты — требует сверки с картой.',
  'Займы — из раздела «Долги» пульта (ведётся вручную), только тип «Займ» / «Учредитель». Долги поставщикам и по зарплате оплачиваются через реестр «Расходы» и уже в строке расходов. «Прочие движения» — разница между изменением остатка и потоками выше: расходы, оплаченные учредителями напрямую (через «Кассу Азиз»), переводы, неразнесённые проводки.',
  'Январь–март 2026 — управленческий пакет (ДДС): январь — выписка банка, февраль–март — поступления Odoo и выплаты по выпискам. Остаток на 01.01.2026 принят 0 (в пакете начальный остаток = займ 1 200 млн, полученный в январе; здесь займ показан строкой «Займы получены»). Остаток на 31.03 — по Odoo, чтобы апрель начинался с факта Odoo; расхождение с пакетом — в «прочих движениях» марта.',
  'Запас = остаток денег на конец месяца ÷ средний отрицательный операционный поток за 3 месяца.'])
widths(ws, {'A': 54, **{L(cc): 15 for cc in range(2, 3 + n)}}); ws.freeze_panes = 'B4'
last3 = [CS[m]['op'] for m in MS[-3:]]; avg3 = sum(last3) / len(last3)
# ---- 3. Выручка ----
ws = wb.create_sheet('Выручка и оплачено вперёд')
title(ws, 'A1', 'MRR, признанная, выставленная и полученная выручка по месяцам (на %s), сум' % CUT_RU)
header(ws, 3, ['Месяц', 'MRR на конец месяца', 'Признано (выручка периода)', 'Выставлено счетов', 'Получено денег', 'Оплачено вперёд за будущие месяцы (на конец месяца)', 'в т.ч. признано SFA', 'в т.ч. признано POS', 'Выставлено SFA', 'Выставлено POS'])
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
  'Скачки «выставлено» по SFA — годовые и полугодовые предоплаты: счёт выставляется разом, а выручка признаётся по месяцам. «Оплачено вперёд» — выставлено до конца месяца за услуги следующих месяцев. По учётной политике компании это уже выручка (в Odoo признаётся сразу), здесь — справочно.',
  'На счёте 201000 в Odoo остался остаток отсрочки по счетам янв–апр 2026 (не списывается с апреля) — по политике «выручка сразу» его нужно сторнировать. Пульт считает по строкам счетов и от этих проводок не зависит.',
  'Дебиторка завышена: в Odoo много поступлений не привязано к счетам (строка выше), поэтому оплаченные счета числятся неоплаченными. Реальную дебиторку покажет сверка платежей со счетами в Odoo.'])
widths(ws, {'A': 46, **{L(cc): 17 for cc in range(2, 11)}}); ws.freeze_panes = 'B4'
# ---- 4. Сегменты ----
ws = wb.create_sheet('Сегменты клиентов')
title(ws, 'A1', 'Сегменты клиентов по размеру: удержание и отток по деньгам (на %s)' % CUT_RU)
header(ws, 3, ['Сегмент', 'Клиентов на срез', 'MRR на срез', 'Доля MRR', 'Лицензий / касс', 'Клиентов на 31.01', 'MRR на 31.01', 'MRR этих клиентов на срез', 'NRR с 31.01', 'GRR с 31.01', 'NRR в пересчёте на год', 'Ушли клиентов с 31.01', 'Потери MRR за 3 мес. (отток + сокращения)'])
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
  'NRR / GRR — по клиентам, у которых был MRR на 31.01.2026: их MRR на срез ÷ MRR на 31.01 (GRR — без учёта роста сверх 31.01). Это %d месяцев; «в пересчёте на год» = NRR^(12/%d). Полные 12 месяцев будут на срез 31.01.2027 — база подписок загружена в журнал MRR Odoo в январе 2026.' % (NMON, NMON),
  'Потери за 3 мес. — все отрицательные движения MRR (отток и сокращения) за последние 3 месяца ÷ MRR сегмента 3 месяца назад.',
  'Когорты — подписки по месяцу создания; М0 = 100%% MRR на конец месяца подключения, дальше — доля этого MRR на конец каждого следующего месяца (с расширениями и сокращениями). Всё по журналу MRR Odoo.'])
widths(ws, {'A': 34, **{L(cc): 15 for cc in range(2, 16)}}); ws.freeze_panes = 'B4'
# ---- 5. NRR и GRR за год ----
ws = wb.create_sheet('NRR и GRR за год')
title(ws, 'A1', 'Удержание выручки (NRR / GRR) за весь доступный период: с 31.01.2026 по %s' % CUT_RU)
header(ws, 3, ['Продукт', 'Клиентов на 31.01', 'MRR на 31.01', 'MRR этих клиентов на срез', 'NRR за %d мес.' % NMON, 'GRR за %d мес.' % NMON, 'NRR в пересчёте на год', 'GRR в пересчёте на год', 'Ушли клиентов'])
r = 4
for lab in ('SFA', 'POS', 'Всего'):
    x = RET[lab]; (totalrow if lab == 'Всего' else datarow)(ws, r, [lab, x['n0'], round(x['mrr0']), round(x['mrr_now']), x['nrr'], x['grr'], x['nrr_y'], x['grr_y'], x['lost']], fmts=[None, NUM, NUM, NUM, PCT, PCT, PCT, PCT, PCT]); r += 1
r += 1; header(ws, r, ['Конец месяца', 'NRR с 31.01 (всего)', 'GRR с 31.01 (всего)', 'NRR SFA', 'GRR SFA', 'NRR POS', 'GRR POS']); r += 1
for i, m in enumerate(MLIST):
    datarow(ws, r, ['%s 2026' % (RUS.get(m) or 'Январь'), RET['Всего']['curve'][i]['nrr'], RET['Всего']['curve'][i]['grr'], RET['SFA']['curve'][i]['nrr'], RET['SFA']['curve'][i]['grr'], RET['POS']['curve'][i]['nrr'], RET['POS']['curve'][i]['grr']], fmts=[None] + [PCT] * 6, blue_inputs=False); r += 1
r += 1; header(ws, r, ['Месяц', 'MRR на начало', 'Новые', 'Расширения всего', 'в т.ч. возврат после сокращения', 'Сокращения всего', 'в т.ч. вернули позже (колебания агентов)', 'Настоящие сокращения', 'Отток', 'GRR месяца', 'GRR месяца без колебаний', 'NRR месяца']); r += 1; r1 = r
for m, x in MONTHLY.items():
    datarow(ws, r, ['%s 2026' % RUS[m], round(x['start']), round(x['new']), round(x['exp']), round(x['exp_back']), round(-x['con']), round(-x['con_back']), round(-x['con_real']), round(-x['churn']), x['grr'], x['grr_adj'], x['nrr']],
            fmts=[None] + [NUM] * 8 + [PCT] * 3, blue_inputs=False); r += 1
totalrow(ws, r, ['ИТОГО'] + [None] + ['=SUM(%s%d:%s%d)' % (L(cc), r1, L(cc), r - 1) for cc in range(3, 10)], fmts=[None] + [NUM] * 8); r += 1
datarow(ws, r, ['GRR за год из месячных: как в отчёте / без колебаний', GRR_REP_Y, GRR_ADJ_Y], fmts=[None, PCT, PCT], blue_inputs=False)
notes(ws, r + 2, [
  'Журнал изменений MRR в Odoo полноценно ведётся с января 2026: тогда в него загружена вся база подписок (483 подписки, ~461 млн MRR). Поэтому 31.01.2026 — самая ранняя точка, от которой можно мерить удержание; полные 12 месяцев будут на срез 31.01.2027.',
  'NRR / GRR за период — по клиентам, у которых был MRR на 31.01.2026: их MRR на срез ÷ MRR на 31.01 (GRR — без роста сверх уровня 31.01). «В пересчёте на год» = значение ^ (12 / %d).' % NMON,
  'Колебания агентов: сокращение, которое позже вернулось расширением той же подписки (клиент временно уменьшил число лицензий и снова увеличил). Настоящие сокращения — то, что не вернулось к срезу. Сокращения последних месяцев ещё могут вернуться.',
  'GRR месяца «как в отчёте» = 1 − (отток + все сокращения) ÷ MRR на начало месяца. «Без колебаний» — в сокращениях только настоящие потери. Перемножать месячный GRR как в отчёте до 12 месяцев некорректно: колебания агентов попадают в потери, хотя возвращаются.',
  'Источник: журнал изменений MRR Odoo (sale.order.log), доллары по 12 000 сум/$; клиент — партнёр подписки; продукт — POS / SFA.'])
widths(ws, {'A': 28, **{L(cc): 17 for cc in range(2, 13)}}); ws.freeze_panes = 'B4'
# ---- 6. История с 2025: одна линия — файл финансиста (2025), пакет (янв–мар 2026), Odoo (с апреля) ----
HM = sorted(HC) + MS; HV = {}
for m in HC:
    h = HC[m]; HV[m] = {'src': h['src'], 'rev': h['rev'], 'exp': h['exp'], 'capex': h.get(FG.CAPEX, 0.0), 'tap': h.get(FG.TAP, 0.0), 'in': h.get('cash', {}).get('in'), 'odoo': None, 'chk': h['chk'].get('ebitda')}
for m in MS:
    HV[m] = {'src': 'odoo', 'rev': {k: REC[(k, m)] for k in FG.REV}, 'exp': {g: EG[(g, m)] for g in FG.GROUPS}, 'capex': CAPEX[m], 'tap': 0.0, 'in': IN[m], 'odoo': INV[m], 'chk': None}
for v in HV.values():
    v['revt'] = sum(v['rev'].values()); v['expt'] = sum(v['exp'].get(g, 0.0) for g in FG.GROUPS); v['ebitda'] = v['revt'] - v['expt']
SRC_RU = {'fin': 'Файл финансиста', 'pack': 'Пакет (PL_Linko)', 'odoo': 'Odoo'}; METH = {'fin': 'по оплатам', 'pack': 'по оплатам', 'odoo': 'по периоду услуги'}
YK = {y: ('Итого %s' % y) if y < '2026' or LAST_M[5:] == '12' else 'Итого %s, янв–%s' % (y, RUS[LAST_M][:3].lower()) for y in sorted({m[:4] for m in HM})}
HROWS = []
def hrow(lab, f, kind='m'):
    vals = {m: f(HV[m]) for m in HM}; tot = {}
    for y, yk in YK.items():
        ms = [m for m in HM if m[:4] == y and vals[m] is not None]
        if kind == 'm' and ms: tot[yk] = sum(vals[m] for m in ms)
        if kind == 'mr' and ms and len(ms) == sum(1 for m in HM if m[:4] == y): tot[yk] = sum(vals[m] for m in ms)   # справочные — итог только за полный ряд
        if kind == 'p' and ms:
            d_ = sum(vals[m][1] for m in ms); tot[yk] = sum(vals[m][0] for m in ms) / d_ if d_ else None
    if kind == 'p': vals = {m: (x[0] / x[1] if x[1] else None) for m, x in vals.items()}
    HROWS.append([lab, kind, vals, tot])
hrow('Источник данных', lambda v: SRC_RU[v['src']], 's'); hrow('Учёт выручки', lambda v: METH[v['src']], 's')
for k in FG.REV: hrow(k, lambda v, k=k: v['rev'].get(k, 0.0))
hrow('ВЫРУЧКА', lambda v: v['revt'])
for g in FG.GROUPS: hrow(g, lambda v, g=g: v['exp'].get(g, 0.0))
hrow('РАСХОДЫ ВСЕГО', lambda v: v['expt']); hrow('EBITDA', lambda v: v['ebitda']); hrow('EBITDA, % выручки', lambda v: (v['ebitda'], v['revt']), 'p')
hrow('Капрасходы: техника и мебель (не в EBITDA)', lambda v: v['capex']); hrow('Финансирование Tap (не в EBITDA)', lambda v: v['tap'])
hrow('Справочно: выручка как в Odoo (по дате счёта)', lambda v: v['odoo'], 'mr'); hrow('Справочно: получено денег от клиентов', lambda v: v['in'], 'mr')
hrow('Справочно: EBITDA в исходном файле', lambda v: v['chk'], 'mr')
if HM:
    ws = wb.create_sheet('История с 2025')
    MRU = lambda m: '%s %s' % (RUS.get(m, ['', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'][int(m[5:])])[:3].lower(), m[:4])
    title(ws, 'A1', 'Выручка, расходы и EBITDA по месяцам с января %s (на %s), сум' % (HM[0][:4], CUT_RU))
    header(ws, 3, ['Статья'] + [MRU(m) for m in HM] + list(YK.values())); r = 4
    for lab, kind, vals, tot in HROWS:
        fmt = PCT if kind == 'p' else None if kind == 's' else NUM; kind = 'm' if kind == 'mr' else kind
        v = [lab] + [(x if kind == 's' else None if x is None else x if kind == 'p' else round(x)) for x in (vals[m] for m in HM)] + [(None if tot.get(yk) is None else tot[yk] if kind == 'p' else round(tot[yk])) for yk in YK.values()]
        (totalrow if lab.isupper() or lab == 'EBITDA' else datarow)(ws, r, v, fmts=[None] + [fmt] * (len(HM) + len(YK))) if lab.isupper() or lab == 'EBITDA' else datarow(ws, r, v, fmts=[None] + [fmt] * (len(HM) + len(YK)), blue_inputs=False)
        r += 1
    notes(ws, r + 1, [
      '2025 — файл финансиста «%s» (лист P&L, факт): учёт по оплатам — выручка = полученные деньги, расходы = оплаты. Техника и мебель и «Финансирование Tap» вынесены под EBITDA, поэтому EBITDA 2025 здесь лучше, чем в файле, ровно на эти суммы (строка «EBITDA в исходном файле»).' % (HIST or {}).get('files', {}).get('fin', 'PL_last_v.xlsx'),
      'Январь–март 2026 — управленческий пакет «%s» (P&L: PL_Linko — файл финансиста, тоже по оплатам). Налоги и сборы / IT Park включены в расходы (в пакете — под EBIT), техника за март — в капрасходы (по ДДС). Биллинг в Odoo до апреля неполный, поэтому выручка февраля–марта — из пакета (решение менеджмента).' % (HIST or {}).get('files', {}).get('pack', 'пакет'),
      'С апреля 2026 — Odoo: выручка по периоду услуги (как на листе «P&L и EBITDA»), расходы — реестр «Расходы» за период + комиссии платёжных систем. Справочно — выручка как в Odoo (по дате счёта, вся сумма сразу) и полученные деньги.',
      'Смена учёта в апреле 2026: до апреля выручка по оплатам (годовая предоплата — в месяце оплаты), с апреля — по периоду услуги (годовая предоплата делится на 12 месяцев). Расходы во все периоды сгруппированы одинаково (единый план статей: fin_groups.py).',
      'Займы и финансирование учредителей — не выручка и не расход: долги ведутся в разделе «Долги», финансирование 2024–2025 — справкой на листе «Баланс». Суммы в долларах — по 12 000 сум/$.'])
    widths(ws, {'A': 46, **{L(cc): 13 for cc in range(2, 3 + len(HM) + len(YK))}}); ws.freeze_panes = 'B4'
# ---- 7. Баланс на срез ----
A_CASH = CASH_END[LAST_M]; GL = F.get('gl', {}); A_AR = GL.get('ar', 0.0); GL201 = GL.get('201000', 0.0)
L_LOAN, L_PAY, L_SUP, L_DEF = DB_BAL['loan'], DB_BAL['pay'], DB_BAL['sup'], DEF.get(LAST_M, 0.0)
A_T = A_CASH + A_AR; L_T = L_LOAN + L_SUP + L_PAY; EQ = A_T - L_T
PBS = (HIST or {}).get('pack_bs', {}); FUND = sum((HIST or {}).get('fund', {}).values()); CHARTER = (PBS.get('BS-301') or {}).get('v')
BAL = [('Денежные средства', A_CASH, 'Odoo: банки, карты, кассы на срез, без «Кассы Азиз» и FOTON; доллары × 12 000 — как на вкладке «Деньги и запас»'),
       ('Дебиторская задолженность клиентов', A_AR, 'Odoo: сальдо счетов расчётов с клиентами на срез = открытые счета минус поступления, ещё не привязанные к счетам'),
       ('Основные средства и НМА', 0.0, 'В Odoo не капитализируются: техника и мебель — капрасходы в месяце покупки, разработка не капитализируется'),
       ('White-label TAP', 0.0, 'Не отражается: компанию TAP забрали целиком, денег по сделке не было ни в приход, ни в расход (решение менеджмента)'),
       ('ИТОГО АКТИВЫ', A_T, ''),
       ('Займы', L_LOAN, 'Раздел «Долги»: тип «Займ» / «Учредитель», остаток на срез (взято − возвращено)'),
       ('Долги поставщикам, аренда и прочие', L_SUP, 'Раздел «Долги»: остальные типы, кроме зарплаты'),
       ('Задолженность по зарплате', L_PAY, 'Раздел «Долги»: тип «Зарплата / сотрудники»'),
       ('ИТОГО ОБЯЗАТЕЛЬСТВА', L_T, ''),
       ('Капитал и накопленный результат', EQ, 'Балансирующая строка: активы − обязательства. Сюда входят уставный капитал и деньги учредителей, не оформленные займами'),
       ('ИТОГО ОБЯЗАТЕЛЬСТВА И КАПИТАЛ', L_T + EQ, '')]
BREF = [('Справочно: оплачено клиентами вперёд за будущие месяцы', L_DEF, 'Расчёт пульта по строкам счетов Odoo: выставлено до среза за услуги после среза. По учётной политике компании это уже выручка (в Odoo признана сразу), поэтому в баланс не входит'),
        ('Справочно: уставный капитал', CHARTER, 'Управленческий баланс пакета на 31.07.2026 (до регистрации увеличения)'),
        ('Справочно: финансирование учредителями, окт 2024 – сен 2025', FUND or None, 'Файл финансиста, лист CF_отчет. Как долг в балансе — только займы из раздела «Долги»'),
        ('Справочно: остаток счёта 201000 в Odoo', -GL201, 'Отсрочка выручки по счетам янв–апр 2026, перестала списываться в апреле. Противоречит политике «выручка сразу» — бухгалтеру сторнировать. Пульт от неё не зависит'),
        ('Справочно: открытые счета клиентов', sum(AR.values()), 'Odoo: неоплаченные счета до среза (вкладка «Выручка»), без зачёта непривязанных поступлений'),
        ('Справочно: поступления, не привязанные к счетам', F['unmatched']['s'], 'Odoo: платежи клиентов без связи со счётом — из-за них открытые счета завышены')]
ws = wb.create_sheet('Баланс на срез')
title(ws, 'A1', 'Управленческий баланс на %s, сум' % CUT_RU); header(ws, 3, ['Статья', 'Сумма, сум', 'Источник и как посчитано']); r = 4
for lab, v, src in BAL + [(None, None, None)] + BREF:
    if lab is None: r += 1; continue
    (totalrow(ws, r, [lab, round(v), src], fmts=[None, NUM]) if lab.startswith('ИТОГО') else datarow(ws, r, [lab, None if v is None else round(v), src], fmts=[None, NUM, None], blue_inputs=False, wrap=True)); r += 1
notes(ws, r + 1, ['Баланс собирается автоматически на каждый срез: деньги и дебиторка — Odoo, все долги — только раздел «Долги» пульта. Капитал — балансирующая строка и может быть отрицательным: компания финансировалась учредителями, а убытки прошлых лет накоплены.',
                  'Отличия от управленческого баланса пакета на 31.07: TAP не включается (в пакете было $350 тыс.), долги — из раздела «Долги» (с выплатами после 31.07).'])
widths(ws, {'A': 50, 'B': 18, 'C': 110}); ws.freeze_panes = 'A4'
# ---- 8. Источники и методика: что откуда и сверки на стыках ----
CHK = []
def chk(what, a_lab, a, b_lab, b, note):
    CHK.append({'what': what, 'a_lab': a_lab, 'a': a, 'b_lab': b_lab, 'b': b, 'diff': (a - b) if a is not None and b is not None else None, 'note': note})
if HC:
    y25 = [m for m in HC if m[:4] == '2025']; q1 = [m for m in HC if m[:4] == '2026']
    if y25: chk('EBITDA 2025', 'Файл финансиста', sum(HC[m]['chk']['ebitda'] for m in y25), 'Пульт (взято)', sum(HV[m]['ebitda'] for m in y25),
                'В пульте техника и мебель (%.0f млн) и «Финансирование Tap» (%.0f млн) вынесены под EBITDA — это не операционные расходы' % (sum(HV[m]['capex'] for m in y25) / 1e6, sum(HV[m]['tap'] for m in y25) / 1e6))
    if q1:
        chk('EBITDA янв–мар 2026', 'Пакет', sum(HC[m]['chk']['ebitda'] for m in q1), 'Пульт (взято)', sum(HV[m]['ebitda'] for m in q1),
            'В пульте налоги и сборы / IT Park — в расходах (как в реестре Odoo с апреля), техника за март — в капрасходах')
        fm_ = [m for m in ('2026-02', '2026-03') if m in HC]
        chk('Выручка февраль–март 2026', 'Пакет, PL_Linko (взято)', sum(HV[m]['revt'] for m in fm_), 'Odoo по периоду услуги', sum(recm(m) for m in fm_),
            'Биллинг в Odoo до апреля неполный (счета заводились не на все оплаты) — взят файл финансиста, решение менеджмента. Odoo по дате счёта: %.0f млн' % (sum(INV[m] for m in fm_) / 1e6))
    if 'cash' in HC.get('2026-03', {}): chk('Остаток денег на 31.03.2026', 'Пакет (ДДС)', HC['2026-03']['cash']['end'], 'Odoo (взято)', CASH_END['2026-03'],
            'Апрель начинается с остатка Odoo; разница — в «прочих движениях» марта на вкладке «Деньги и запас»')
    pk = (HIST or {}).get('pack_chk', {}); a4 = [m for m in ('2026-04', '2026-05', '2026-06', '2026-07') if m in MS and m in pk]
    if a4: chk('Выручка апрель–%s 2026' % RUS[a4[-1]].lower(), 'Пакет: оплаченные счета Odoo', sum(pk[m]['rev'] for m in a4), 'Пульт: по периоду услуги (взято)', sum(recm(m) for m in a4),
               'Пакет брал только оплаченные счета по дате; пульт — все счета, разнесённые по периоду услуги. Как в Odoo (все счета по дате): %.0f млн' % (sum(INV[m] for m in a4) / 1e6))
chk('Оплачено вперёд за будущие месяцы на срез', 'Пульт: по строкам счетов (справочно)', L_DEF, 'Odoo: остаток счёта 201000', -GL201,
    'Odoo формировал отсрочку только по счетам янв–апр 2026 и перестал списывать её в апреле; с мая по политике компании выручка признаётся сразу. Пульт считает заново по периодам в строках счетов, в баланс не включает')
chk('Дебиторка на срез', 'Открытые счета клиентов', sum(AR.values()), 'Сальдо счетов клиентов Odoo (в балансе)', A_AR,
    'Разница — поступления, не привязанные к счетам (%.0f млн): оплаченные счета числятся открытыми. Нужна привязка платежей в Odoo' % (F['unmatched']['s'] / 1e6))
if LOANS_OK and '2026-07' in DEBT_END and PBS: chk('Займы на 31.07.2026', 'Раздел «Долги» (взято)', DEBT_END['2026-07'], 'Пакет, баланс 31.07 (BS-240…244)', sum((PBS.get(k) or {}).get('v', 0) for k in ('BS-240', 'BS-241', 'BS-242', 'BS-243', 'BS-244')),
                 'Все долги ведутся только в разделе «Долги»; на срез займы — %.0f млн (после 31.07 добавлены новые займы и выплаты)' % (L_LOAN / 1e6))
SRCS = [('Янв–дек 2025', 'Файл финансиста PL_last_v.xlsx, лист P&L (факт)', 'По оплатам', 'Выручка = деньги от клиентов, расходы = оплаты. Техника и «Финансирование Tap» вынесены под EBITDA'),
        ('Янв–мар 2026', 'Управленческий пакет: P&L (PL_Linko) и ДДС', 'По оплатам', 'Выручка — файл финансиста (в Odoo биллинг до апреля неполный). Деньги: январь — выписка, февраль–март — Odoo и выписки'),
        ('С апреля 2026', 'Odoo linktrade (только чтение) + раздел «Долги» пульта', 'По периоду услуги', 'Выручка — строки счетов по дням периода; расходы — реестр «Расходы» за период; деньги — остатки счетов Odoo'),
        ('Долги', 'Раздел «Долги» пульта (ведётся вручную)', '—', 'Единственный источник по займам, долгам поставщикам и по зарплате; из Odoo долги не берутся'),
        ('Курс', 'Управленческий курс', '12 000 сум/$', 'Все суммы в долларах (подписки, расходы, валютные счета, долги)')]
ws = wb.create_sheet('Источники и методика')
title(ws, 'A1', 'Источники данных и методика (на %s)' % CUT_RU); header(ws, 3, ['Период / блок', 'Источник', 'Учёт выручки', 'Как посчитано']); r = 4
for row in SRCS: datarow(ws, r, list(row), blue_inputs=False, wrap=True); r += 1
r += 1; header(ws, r, ['Сверка на стыке', 'Источник А', 'Сумма А', 'Источник Б', 'Сумма Б', 'Разница', 'Что взято и почему']); r += 1
for x in CHK: datarow(ws, r, [x['what'], x['a_lab'], None if x['a'] is None else round(x['a']), x['b_lab'], None if x['b'] is None else round(x['b']), None if x['diff'] is None else round(x['diff']), x['note']], fmts=[None, None, NUM, None, NUM, NUM, None], blue_inputs=False, wrap=True); r += 1
widths(ws, {'A': 30, 'B': 34, 'C': 16, 'D': 34, 'E': 16, 'F': 14, 'G': 90})
wb._sheets = [ws2 for nm in ('P&L и EBITDA', 'История с 2025', 'Деньги и запас', 'Баланс на срез', 'Выручка и оплачено вперёд', 'Сегменты клиентов', 'NRR и GRR за год', 'Источники и методика') for ws2 in wb.worksheets if ws2.title == nm]
fn = OUTP('Финансы_для_банка_%s_2026_%s.xlsx' % (PERIOD.replace('–', '-'), SUF)); wb.save(fn)
res = {'MC': MC, 'BR': BR, 'HIST': {'M': HM, 'ROWS': HROWS, 'Y': list(YK.values())}, 'BAL': BAL, 'BREF': BREF, 'CHK': CHK, 'SRCS': SRCS, 'HFILES': (HIST or {}).get('files'),
       'CAPEX': {m: CAPEX[m] for m in MS}, 'MS': MS, 'MR': MR, 'PL': PL, 'REC': {'%s|%s' % kk: v for kk, v in REC.items() if kk[1] in MR}, 'INV': {m: INV[m] for m in MR}, 'IN': {m: IN[m] for m in MR}, 'MRR': MRR_END,
       'DEF': {m: DEF.get(m, 0) for m in MR}, 'OPX': {lab: v for lab, v in OPX}, 'COGS_FOT': COGS_FOT, 'HOST': HOST, 'CS': CS, 'AVG3': avg3, 'LOANS_OK': LOANS_OK,
       'ACC': {('%s · %s' % (code, ACC[code]['name'])): bal_at(code, LAST_M) for code in ACC}, 'AR': {b: AR[b] for b in BK}, 'ARN': {b: ARN[b] for b in BK}, 'UNM': F['unmatched'],
       'SEG': SR, 'SEGS': SEGS, 'COH': COH, 'NMON': NMON, 'RET': RET, 'MONTHLY': MONTHLY, 'GRR_ADJ_Y': GRR_ADJ_Y, 'GRR_REP_Y': GRR_REP_Y, 'CON_T': CON_T, 'CON_B': CON_B}
json.dump(res, open(J('res_fin'), 'w'), ensure_ascii=False)
print(fn)
for m in MS: print(m, {k: round(v / 1e6, 1) for k, v in PL[m].items()}, '| деньги', {k: round((v or 0) / 1e6, 1) for k, v in CS[m].items() if not isinstance(v, str)})
print('avg3', round(avg3 / 1e6, 1), 'loans', LOANS_OK)
for m in MC[:3]: print(m, {k: (v if isinstance(v, str) else round((v or 0) / 1e6, 1)) for k, v in CS[m].items()})
for m in MS: print('мостик', m, {k: round(v / 1e6, 1) for k, v in BR[m].items()}, 'odoo', round(INV[m] / 1e6, 1), 'ebitda_odoo', round(PL[m]['ebitda_odoo'] / 1e6, 1))
for lab, v, _ in BAL + BREF: print('БАЛ', lab, None if v is None else round(v / 1e6, 1))
for x in CHK: print('СВЕРКА', x['what'], x['a_lab'], round((x['a'] or 0) / 1e6, 1), '|', x['b_lab'], round((x['b'] or 0) / 1e6, 1))
for lab, kind, vals, tot in HROWS: print('ИСТ', lab, {k: (round(v / 1e6, 1) if kind == 'm' else round(v, 3) if v is not None else None) for k, v in tot.items()})
for m in MR: print(m, 'MRR', round(MRR_END[m] / 1e6, 1), 'rec', round(recm(m) / 1e6, 1), 'inv', round(INV[m] / 1e6, 1), 'in', round(IN[m] / 1e6, 1), 'def', round(DEF.get(m, 0) / 1e6, 1))
for s, v in SR.items(): print(s, {k: (round(x, 3) if isinstance(x, float) and abs(x) < 10 else round(x / 1e6, 1) if isinstance(x, float) else x) for k, x in v.items()})
print('AR', {b: round(AR[b] / 1e6, 1) for b in BK}, F['unmatched'])
for m, cdata in COH.items(): print(m, cdata['n'], round(cdata['start'] / 1e6, 1), [round(x, 2) for x in cdata['k']])
print('RET', {k: {kk: (round(v, 3) if isinstance(v, float) and abs(v) < 10 else v) for kk, v in x.items() if kk != 'curve'} for k, x in RET.items()})
for m, x in MONTHLY.items(): print(m, {k: (round(v / 1e6, 1) if abs(v) > 10 else round(v, 3)) for k, v in x.items() if v is not None})
print('GRR год как в отчёте', GRR_REP_Y, 'без колебаний', GRR_ADJ_Y, 'сокращения вернулись', CON_B / CON_T if CON_T else None)
