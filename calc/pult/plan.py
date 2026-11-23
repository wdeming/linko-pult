# -*- coding: utf-8 -*-
"""План поступлений денег на 6 месяцев вперёд от сегодняшнего дня (только чтение Odoo).
python3 pult/plan.py <папка вывода>  → plan.json и План_поступлений_на_ДД.ММ.ГГГГ.xlsx
1) Подписки POS / SFA в работе и на паузе: оплачено по (конец периода последнего счёта, следующая дата счёта в подписке), дальше
   каждый следующий период подписки (по её плану — месяц, 2–12 мес.) — ожидаемая оплата суммы периода в дату начала периода.
   Неоплаченные периоды, начавшиеся до сегодня, — «просрочено» (в план месяцев не входит).
2) Поправка на отток — средний месячный отток MRR за последние 6 полных месяцев (журнал MRR), нарастающим итогом.
3) Новые продажи и разовые (обучение, доработки, Perfect Store, производство) — среднее за последние 6 полных месяцев по счетам."""
import sys, os, json, datetime, calendar, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__))); sys.path.insert(1, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from odoo_db import OdooDB
from cls import PROD_SQL
K = 12000; OUT = sys.argv[1] if len(sys.argv) > 1 else '.'
TODAY = datetime.date.fromisoformat(os.environ['PLAN_TODAY']) if os.environ.get('PLAN_TODAY') else datetime.date.today()
def addm(d, n):
    y, m = divmod(d.month - 1 + n, 12); y += d.year; m += 1
    return datetime.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))
ME = lambda d: d.replace(day=calendar.monthrange(d.year, d.month)[1])
M0 = TODAY.replace(day=1); CUR_END = ME(TODAY)
HM = [addm(M0, i + 1).strftime('%Y-%m') for i in range(6)]                 # 6 полных месяцев после текущего
H_END = ME(datetime.date.fromisoformat(HM[-1] + '-01'))
RM = ['', 'янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']
ml = lambda m: '%s %s' % (RM[int(m[5:7])], m[2:4])
C_OVER, C_CUR = 'Просрочено на сегодня', '%s %s (до %s)' % (RM[TODAY.month], str(TODAY.year)[2:], CUR_END.strftime('%d.%m'))
COLS = [C_OVER, C_CUR] + [ml(m) for m in HM]
db = OdooDB(); q = db.query
# ---------------- 1. подписки ----------------
SUBS = q(f"""select so.id, so.name, {PROD_SQL} prod, so.subscription_state st, so.next_invoice_date nid, so.start_date sd, so.currency_id cur,
  so.recurring_total rt, so.recurring_monthly rm, coalesce(p.billing_period_value, 1) bp, coalesce(pl.name->>'ru_RU', pl.name->>'en_US') plan, rp.name client,
  (select max(aml.deferred_end_date) from sale_order_line sol join sale_order_line_invoice_rel rel on rel.order_line_id=sol.id join account_move_line aml on aml.id=rel.invoice_line_id
     join account_move am on am.id=aml.move_id where sol.order_id=so.id and am.state='posted' and am.move_type='out_invoice') cov,
  (select max(am.invoice_date) from sale_order_line sol join sale_order_line_invoice_rel rel on rel.order_line_id=sol.id join account_move_line aml on aml.id=rel.invoice_line_id
     join account_move am on am.id=aml.move_id where sol.order_id=so.id and am.state='posted' and am.move_type='out_invoice') linv
 from sale_order so left join sale_subscription_plan p on p.id=so.plan_id left join sale_subscription_plan pl on pl.id=so.plan_id left join res_partner rp on rp.id=so.partner_id
 where so.is_subscription and so.subscription_state in ('3_progress','4_paused')""")
ST = {'3_progress': 'в работе', '4_paused': 'на паузе'}
LINE = lambda p, st: 'Подписки %s — %s' % (p, ST[st])
SCH = collections.defaultdict(float); OVER = collections.defaultdict(float); BYLEN = collections.defaultdict(float); DET = []; AGE = collections.defaultdict(lambda: [0, 0.0])
plen = lambda bp: 'Помесячно' if bp <= 1 else '2–3 месяца' if bp <= 3 else '4–6 месяцев' if bp <= 6 else '7–12 месяцев'
for s in SUBS:
    if s['prod'] not in ('POS', 'SFA'): continue
    bp = int(s['bp'] or 1); amt = float(s['rt'] or 0) * (K if s['cur'] == 1 else 1); mrr = float(s['rm'] or 0) * (K if s['cur'] == 1 else 1)
    c = [x for x in (s['cov'], (s['nid'] - datetime.timedelta(days=1)) if s['nid'] else None) if x]
    if c: pt = max(c); basis = 'период последнего счёта / следующая дата счёта в подписке'
    elif s['linv']: pt = addm(s['linv'], bp) - datetime.timedelta(days=1); basis = 'дата последнего счёта + период плана'
    else: pt = (s['sd'] or TODAY) - datetime.timedelta(days=1); basis = 'счетов нет — с даты начала подписки'
    line = LINE(s['prod'], s['st']); d = s['prod'] + '|' + s['st']
    row = {'name': s['name'], 'client': s['client'], 'prod': s['prod'], 'st': ST[s['st']], 'plan': s['plan'] or '%d мес.' % bp, 'bp': bp, 'amt': amt, 'mrr': mrr, 'pt': pt.isoformat(), 'basis': basis, 'over': 0.0, 'over_n': 0, 'm': {}}
    nxt = pt + datetime.timedelta(days=1); first_over = None
    while nxt <= H_END and amt > 0:
        if nxt < TODAY:
            OVER[line] += amt; row['over'] += amt; row['over_n'] += 1; first_over = first_over or nxt
        else:
            col = C_CUR if nxt <= CUR_END else ml(nxt.strftime('%Y-%m'))
            SCH[(line, col)] += amt; BYLEN[(plen(bp), col)] += amt; row['m'][col] = row['m'].get(col, 0) + amt
            row.setdefault('next', nxt.isoformat())
        nxt = addm(nxt, bp)
    if first_over:
        dd = (TODAY - first_over).days; b = 'до 30 дней' if dd <= 30 else '31–90 дней' if dd <= 90 else '91–180 дней' if dd <= 180 else 'больше 180 дней'
        AGE[(s['prod'] + ' ' + ST[s['st']], b)][0] += 1; AGE[(s['prod'] + ' ' + ST[s['st']], b)][1] += row['over']; row['age'] = b
    DET.append(row)
# ---------------- 2. отток: средний месячный за 6 полных месяцев ----------------
P6 = [addm(M0, -i).strftime('%Y-%m') for i in range(6, 0, -1)]
LOG = q(f"""select {PROD_SQL} prod, to_char(l.event_date,'YYYY-MM') m, l.event_type et, sum(case when l.currency_id=1 then l.amount_signed*{K} else l.amount_signed end) a
  from sale_order_log l join sale_order so on so.id=l.order_id where l.event_date < date '{M0}' group by 1,2,3""")
mv = collections.defaultdict(float)
for r in LOG: mv[(r['prod'], r['m'], r['et'])] += float(r['a'])
CHURN = {}
for p in ('SFA', 'POS'):
    num = den = 0.0
    for m in P6:
        start = sum(v for (pp, mm, et), v in mv.items() if pp == p and mm < m)
        num += -mv[(p, m, '2_churn')]; den += start
    CHURN[p] = num / den if den else 0.0
# ---------------- 3. новые продажи и разовые: среднее за 6 полных месяцев по счетам ----------------
SFA_P, POS_P = (25, 26, 41, 42), (24, 40, 43, 32)
OT = {30: 'Обучение', 33: 'Обучение', 37: 'Обучение', 38: 'Доработки', 39: 'Доработки'}
INVL = q(f"""select l.product_id pid, m.invoice_date d, so.id so_id,
   (case when m.move_type='out_refund' then -1 else 1 end)*(case when m.currency_id=1 then l.price_subtotal*{K} else l.price_subtotal end) a
 from account_move_line l join account_move m on m.id=l.move_id
 left join sale_order_line_invoice_rel rel on rel.invoice_line_id=l.id left join sale_order_line sol on sol.id=rel.order_line_id left join sale_order so on so.id=sol.order_id and so.is_subscription
 where m.state='posted' and m.move_type in ('out_invoice','out_refund') and l.display_type='product'""")
first = {}
for x in INVL:
    if x['so_id'] and x['pid'] in SFA_P + POS_P: first[x['so_id']] = min(first.get(x['so_id'], x['d']), x['d'])
NEWF = collections.defaultdict(float); ONE = collections.defaultdict(float)
for x in INVL:
    m = x['d'].strftime('%Y-%m')
    if m not in P6: continue
    a = float(x['a'] or 0)
    if x['pid'] in SFA_P + POS_P:
        if x['so_id'] and first.get(x['so_id']) and first[x['so_id']].strftime('%Y-%m') == m: NEWF[('SFA' if x['pid'] in SFA_P else 'POS', m)] += a
    else: ONE[(OT.get(x['pid'], 'Perfect Store, производство и прочее'), m)] += a
NEWM = {}      # новый MRR помесячных подписок — они платят и в следующие месяцы горизонта
for p in ('SFA', 'POS'):
    NEWM[p] = sum(float(r['rm'] or 0) * (K if r['cur'] == 1 else 1) for r in q(f"""select so.recurring_monthly rm, so.currency_id cur, {PROD_SQL} prod from sale_order so
        where so.is_subscription and coalesce(so.plan_id,1)=1 and so.first_contract_date >= date '{P6[0]}-01' and so.first_contract_date < date '{M0}'""") if r['prod'] == p) / 6
AVG_NEW = {p: sum(NEWF[(p, m)] for m in P6) / 6 for p in ('SFA', 'POS')}
ONE_G = ['Обучение', 'Доработки', 'Perfect Store, производство и прочее']
AVG_ONE = {g: sum(ONE[(g, m)] for m in P6) / 6 for g in ONE_G}
IN6 = q(f"""select coalesce(sum(case when c.name='UZS' then p.amount_company_currency_signed else p.amount*{K} end),0) s from account_payment p join account_move am on am.id=p.move_id
  join res_currency c on c.id=p.currency_id where am.state='posted' and p.payment_type='inbound' and not p.is_internal_transfer and am.date >= date '{P6[0]}-01' and am.date < date '{M0}'""")[0]['s']
AVG_IN = float(IN6) / 6
# ---------------- свод ----------------
REST = ((CUR_END - TODAY).days + 1) / calendar.monthrange(TODAY.year, TODAY.month)[1]    # доля оставшегося текущего месяца
KI = {C_CUR: 0}; KI.update({ml(m): i + 1 for i, m in enumerate(HM)}); SH = {C_CUR: REST}; SH.update({ml(m): 1.0 for m in HM})
FC = COLS[1:]
TAB = collections.OrderedDict()
SUBL = [LINE(p, st) for p in ('SFA', 'POS') for st in ('3_progress', '4_paused')]
for l in SUBL: TAB[l] = {C_OVER: OVER[l], **{c: SCH[(l, c)] for c in FC}}
TAB['ПОДПИСКИ ПО ГРАФИКУ'] = {c: sum(TAB[l][c] for l in SUBL) for c in COLS}
adj = {}
for c in FC:
    adj[c] = -sum(SCH[(LINE(p, st), c)] * (1 - (1 - CHURN[p]) ** KI[c]) for p in ('SFA', 'POS') for st in ('3_progress', '4_paused'))
TAB['Поправка на отток'] = adj
TAB['ПОДПИСКИ С УЧЁТОМ ОТТОКА'] = {c: TAB['ПОДПИСКИ ПО ГРАФИКУ'][c] + adj[c] for c in FC}
for p in ('SFA', 'POS'):
    TAB['Новые продажи %s (прогноз)' % p] = {c: AVG_NEW[p] * SH[c] + NEWM[p] * max(0, KI[c] - 1 + (0 if c == C_CUR else REST)) for c in FC}
for g in ONE_G: TAB['%s (прогноз)' % g] = {c: AVG_ONE[g] * SH[c] for c in FC}
TAB['ВСЕГО ПЛАН ПОСТУПЛЕНИЙ'] = {c: TAB['ПОДПИСКИ С УЧЁТОМ ОТТОКА'][c] + sum(TAB[k][c] for k in TAB if k.endswith('(прогноз)')) for c in FC}
TAB['Справочно: среднее фактическое поступление за 6 мес.'] = {c: AVG_IN * SH[c] for c in FC}
for k, v in TAB.items(): v['Итого за горизонт'] = sum(v.get(c, 0) or 0 for c in FC)
BIG = sorted([r for r in DET if r['m']], key=lambda r: -max(r['m'].values()))[:40]
res = {'today': TODAY.isoformat(), 'cols': COLS, 'months': HM, 'tab': list(TAB.items()),   # список — порядок строк сохраняется в JSON 'bylen': {k[0] + '|' + k[1]: v for k, v in BYLEN.items()},
       'age': [[k[0], k[1], v[0], v[1]] for k, v in sorted(AGE.items())], 'churn': CHURN, 'avg_new': AVG_NEW, 'new_mrr': NEWM, 'avg_one': AVG_ONE, 'avg_in': AVG_IN, 'p6': P6,
       'n_subs': {LINE(p, st): sum(1 for r in DET if r['prod'] == p and r['st'] == ST[st]) for p in ('SFA', 'POS') for st in ('3_progress', '4_paused')},
       'big': [{'name': r['name'], 'client': r['client'], 'prod': r['prod'], 'st': r['st'], 'plan': r['plan'], 'amt': r['amt'], 'next': r.get('next'), 'pt': r['pt']} for r in BIG]}
# ---------------- Excel ----------------
import openpyxl
from openpyxl.utils import get_column_letter as LL
from style_lib import *
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'План поступлений'
title(ws, 'A1', 'План поступлений денег на %s – %s (на %s), сум' % (ml(HM[0]), ml(HM[-1]), TODAY.strftime('%d.%m.%Y')))
header(ws, 3, ['Статья'] + COLS + ['Итого за горизонт']); r = 4
for k, v in TAB.items():
    vals = [k] + [None if v.get(c) is None else round(v[c]) for c in COLS] + [round(v['Итого за горизонт'])]
    (totalrow(ws, r, vals, fmts=[None] + [NUM] * (len(COLS) + 1)) if k.isupper() else datarow(ws, r, vals, fmts=[None] + [NUM] * (len(COLS) + 1), blue_inputs=False)); r += 1
notes(ws, r + 1, [
  'Подписки — все подписки POS и SFA в работе и на паузе. «Оплачено по» = конец периода последнего счёта (или следующая дата счёта в подписке). Дальше каждый следующий период по плану подписки (месяц, 3, 6, 12 мес.) — ожидаемая оплата суммы периода в дату его начала. Сумма — текущая сумма подписки за период (доллары × 12 000).',
  '«Просрочено на сегодня» — периоды, которые начались до %s и не оплачены (счёта нет). В итог горизонта не входит: сколько из этого реально соберут — отдельный вопрос (см. лист «Просрочено»).' % TODAY.strftime('%d.%m.%Y'),
  'Поправка на отток — средний месячный отток MRR за %s – %s по журналу MRR Odoo: SFA %.1f%%, POS %.1f%% в месяц, нарастающим итогом по месяцам горизонта.' % (ml(P6[0]), ml(P6[-1]), CHURN['SFA'] * 100, CHURN['POS'] * 100),
  'Новые продажи — средняя первая оплата новых подписок за 6 мес. (SFA %.0f млн, POS %.0f млн в месяц) + ежемесячные оплаты новых помесячных подписок, подключённых в горизонте (новый MRR помесячных: SFA %.1f млн, POS %.1f млн в месяц).' % (AVG_NEW['SFA'] / 1e6, AVG_NEW['POS'] / 1e6, NEWM['SFA'] / 1e6, NEWM['POS'] / 1e6),
  'Разовые — среднее по счетам за 6 мес.: обучение %.1f млн, доработки %.1f млн, Perfect Store / производство / прочее %.1f млн в месяц. Текущий месяц — пропорционально оставшимся дням.' % tuple(AVG_ONE[g] / 1e6 for g in ONE_G),
  'Не учтено: расширения у текущих клиентов, повышение цен с 2027 года, оплаты долгов прошлых периодов.'])
widths(ws, {'A': 50, **{LL(c): 15 for c in range(2, len(COLS) + 3)}}); ws.freeze_panes = 'B4'
ws = wb.create_sheet('По подпискам')
header(ws, 1, ['Подписка', 'Клиент', 'Продукт', 'Статус', 'План оплаты', 'Сумма за период', 'MRR', 'Оплачено по', 'Как определено', 'Следующая оплата', 'Просрочено, сумма', 'Просрочено, периодов'] + FC); r = 2
for x in sorted(DET, key=lambda x: (x['prod'], x['st'], x['client'] or '')):
    datarow(ws, r, [x['name'], x['client'], x['prod'], x['st'], x['plan'], round(x['amt']), round(x['mrr']), x['pt'], x['basis'], x.get('next'), round(x['over']) or None, x['over_n'] or None] + [round(x['m'][c]) if c in x['m'] else None for c in FC],
            fmts=[None] * 5 + [NUM, NUM] + [None] * 3 + [NUM, NUM] + [NUM] * len(FC), blue_inputs=False); r += 1
widths(ws, {'A': 10, 'B': 36, 'C': 7, 'D': 10, 'E': 13, 'F': 14, 'G': 12, 'H': 12, 'I': 30, 'J': 13, 'K': 14, 'L': 10, **{LL(c): 13 for c in range(13, 13 + len(FC))}}); ws.freeze_panes = 'C2'; ws.auto_filter.ref = 'A1:%s%d' % (LL(12 + len(FC)), r - 1)
ws = wb.create_sheet('Просрочено')
header(ws, 1, ['Продукт и статус', 'Срок с первого неоплаченного периода', 'Подписок', 'Сумма, сум']); r = 2
for k0, k1, n, a in res['age']: datarow(ws, r, [k0, k1, n, round(a)], fmts=[None, None, NUM, NUM], blue_inputs=False); r += 1
widths(ws, {'A': 22, 'B': 34, 'C': 10, 'D': 16})
fn = 'План_поступлений_на_%s.xlsx' % TODAY.strftime('%d.%m.%Y'); wb.save(os.path.join(OUT, fn)); res['file'] = fn
json.dump(res, open(os.path.join(OUT, 'plan.json'), 'w'), ensure_ascii=False, default=str)
print(json.dumps({'ok': True, 'file': fn}))
