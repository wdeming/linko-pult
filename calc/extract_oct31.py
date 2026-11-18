# -*- coding: utf-8 -*-
"""Единая выгрузка для комплекта на срез cfg_cut.CUT: база подписок, движения MRR, когорты, строки, мета,
выживаемость, константы карточки, хостинг, штат отделов себестоимости, затраты на привлечение, каналы."""
import json, collections, datetime as dt, re
from odoo_db import OdooDB
from cls import PROD_SQL, MRR_SQL, AMT_SQL
from cfg_cut import *
db = OdooDB()
def dump(name, obj): json.dump(obj, open(J(name), 'w'), ensure_ascii=False, default=str)

# ───────── 1. база подписок на срез ─────────
q = f"""
with ev as (
  select order_id,
    max(case when event_type='0_creation' and event_date > date '{CUT}' then 1 else 0 end) created_after,
    max(case when event_type='2_churn'    and event_date > date '{CUT}' then 1 else 0 end) churned_after,
    coalesce(sum(case when event_date > date '{CUT}'
      then case when currency_id=1 then amount_signed*12000 else amount_signed end else 0 end),0) amt_after
  from sale_order_log group by order_id)
select so.id, so.name, so.partner_id, rp.name pname, {PROD_SQL} prod, so.subscription_state st_now,
  case when so.subscription_state in ('3_progress','4_paused') then {MRR_SQL} else 0 end mrr_now,
  coalesce(ev.amt_after,0) amt_after, so.currency_id, so.plan_id, sp.billing_period_value pv, sp.billing_period_unit pu,
  so.end_date, so.next_invoice_date, so.first_contract_date, so.pricelist_id, pl.name plname,
  (select string_agg(sol.product_id::text||':'||sol.product_uom_qty::text, '|' order by sol.product_id)
     from sale_order_line sol where sol.order_id=so.id and sol.product_id is not null) lines
from sale_order so
left join res_partner rp on rp.id=so.partner_id
left join ev on ev.order_id=so.id
left join sale_subscription_plan sp on sp.id=so.plan_id
left join product_pricelist pl on pl.id=so.pricelist_id
where coalesce(ev.created_after,0)=0
  and ((so.subscription_state in ('3_progress','4_paused')) or coalesce(ev.churned_after,0)=1)"""
rows = db.query(q)
# статус на срез — по истории статусов (mail.tracking): старое значение первой смены статуса после среза
MAP = {'In Progress':'3_progress','В процессе':'3_progress','Paused':'4_paused','Приостановлено':'4_paused',
       'Churned':'6_churn','Отток':'6_churn'}
trk = db.query(f"""select distinct on (m.res_id) m.res_id oid, tv.old_value_char o
 from mail_tracking_value tv join ir_model_fields f on f.id=tv.field_id join mail_message m on m.id=tv.mail_message_id
 where m.model='sale.order' and f.name='subscription_state' and m.date > timestamp '{CUT_UTC}'
 order by m.res_id, m.date, tv.id""")
first_after = {r['oid']: MAP.get(r['o']) for r in trk}
base = []; healed = 0
for r in rows:
    st = first_after.get(r['id']) or r['st_now']
    if st == '6_churn':          # на срез была закрыта, после среза возвращена в работу — в базе, считаем действующей
        st = '3_progress'; healed += 1
    if st not in ('3_progress', '4_paused'): st = '3_progress'
    pn = r['plname']
    base.append({'id': r['id'], 'name': r['name'], 'partner_id': r['partner_id'], 'pname': r['pname'],
        'prod': r['prod'], 'st': st, 'mrr': float(r['mrr_now']) - float(r['amt_after']), 'cur': r['currency_id'],
        'plan_id': r['plan_id'], 'pv': r['pv'], 'pu': r['pu'],
        'end_date': str(r['end_date']) if r['end_date'] else None,
        'next_invoice': str(r['next_invoice_date']) if r['next_invoice_date'] else None,
        'first_contract': str(r['first_contract_date']) if r['first_contract_date'] else None,
        'plname': (pn.get('ru_RU') or pn.get('en_US')) if isinstance(pn, dict) else pn, 'lines': r['lines']})
dump('base', base)
agg = collections.defaultdict(lambda: [0, 0.0])
for b in base: agg[(b['prod'], b['st'])][0] += 1; agg[(b['prod'], b['st'])][1] += b['mrr']
print('подписок в срезе:', len(base), '| возвращённых из оттока после среза:', healed)
for k in sorted(agg): print('  ', k, agg[k][0], int(agg[k][1]))
ids = [b['id'] for b in base]

# ───────── 2. строки, мета ─────────
lines = db.query("""select sol.order_id oid, sol.product_id pid, sol.product_uom_qty::float qty,
  (case when so.currency_id=1 then sol.price_subtotal*12000 else sol.price_subtotal end)::float amt
 from sale_order_line sol join sale_order so on so.id=sol.order_id
 where sol.product_id is not null and sol.order_id = any(%s)""", (ids,))
dump('lines', lines)
meta = db.query("""select so.id, ct.name->>'ru_RU' team, pl.name->>'ru_RU' pricelist, sp.name->>'ru_RU' plan
 from sale_order so left join crm_team ct on ct.id=so.team_id
 left join product_pricelist pl on pl.id=so.pricelist_id left join sale_subscription_plan sp on sp.id=so.plan_id
 where so.id = any(%s)""", (ids,))
dump('meta', {str(r['id']): {'team': r['team'], 'pricelist': r['pricelist'], 'plan': r['plan']} for r in meta})

# ───────── 3. движения MRR и когорты ─────────
out = {}
mov = db.query(f"""select to_char(l.event_date,'YYYY-MM') m, {PROD_SQL} prod, l.event_type et, sum({AMT_SQL}) amt, count(*) n
 from sale_order_log l join sale_order so on so.id=l.order_id
 where l.event_date>=date '{P1}' and l.event_date<=date '{CUT}' group by 1,2,3""")
out['moves'] = [{'m': r['m'], 'prod': r['prod'], 'et': r['et'], 'amt': float(r['amt']), 'n': r['n']} for r in mov]
coh = db.query(f"""
with s as (select so.id, {PROD_SQL} prod, to_char(so.first_contract_date,'YYYY-MM') coh,
    (select min(l.event_date) from sale_order_log l where l.order_id=so.id and l.event_type='2_churn') churn_date
  from sale_order so where so.first_contract_date >= date '{P1}' and so.first_contract_date <= date '{CUT}'
    and exists (select 1 from sale_order_log l where l.order_id=so.id))
select coh, prod, count(*) n, count(*) filter (where churn_date is null or churn_date > date '{CUT}') alive
from s group by 1,2 order by 1,2""")
out['cohorts'] = [dict(r) for r in coh]
cev = db.query(f"""select to_char(so.first_contract_date,'YYYY-MM') coh, {PROD_SQL} prod, l.event_type et, sum({AMT_SQL}) amt
 from sale_order_log l join sale_order so on so.id=l.order_id
 where so.first_contract_date >= date '{P1}' and so.first_contract_date <= date '{CUT}' and l.event_date <= date '{CUT}'
 group by 1,2,3 order by 1,2,3""")
out['cohort_events'] = [{'coh': r['coh'], 'prod': r['prod'], 'et': r['et'], 'amt': float(r['amt'])} for r in cev]
cf = db.query(f"""
with s as (select so.id, {PROD_SQL} prod, to_char(so.first_contract_date,'YYYY-MM') coh,
   case when so.subscription_state in ('3_progress','4_paused') then {MRR_SQL} else 0 end mrr_now
   from sale_order so where so.first_contract_date >= date '{P1}' and so.first_contract_date <= date '{CUT}')
select s.coh, s.prod, sum(s.mrr_now) mrr_now,
  coalesce((select sum(case when l.currency_id=1 then l.amount_signed*12000 else l.amount_signed end)
     from sale_order_log l where l.order_id in (select id from s s2 where s2.coh=s.coh and s2.prod=s.prod)
       and l.event_date > date '{CUT}'),0) aft
from s group by 1,2 order by 1,2""")
out['cohort_fact'] = [{'coh': r['coh'], 'prod': r['prod'], 'mrr_now': float(r['mrr_now']), 'aft': float(r['aft'])} for r in cf]
# количество событий по месяцам (для листов оттока и карточки)
out['events_n'] = [dict(r) for r in db.query(f"""select to_char(l.event_date,'YYYY-MM') m, {PROD_SQL} prod, l.event_type et, count(*) n
 from sale_order_log l join sale_order so on so.id=l.order_id
 where l.event_date>=date '{P1}' and l.event_date<=date '{CUT}' group by 1,2,3""")]
dump('mrr', out)

# ───────── 4. когортная выживаемость ─────────
rs_all = db.query(f"""select {PROD_SQL} prod, date_trunc('month', so.first_contract_date)::date coh, so.id,
  (select min(l.event_date) from sale_order_log l where l.order_id=so.id and l.event_type='2_churn' and l.event_date<=date '{CUT}') churn_date
 from sale_order so where so.is_subscription and so.first_contract_date >= date '{P1}' and so.first_contract_date <= date '{CUT}'""")
surv = {}
for prod in ('SFA', 'POS'):
    rs = [r for r in rs_all if r['prod'] == prod]; o = []; n = 0
    while True:
        den = [r for r in rs if (CUT_D.year - r['coh'].year) * 12 + (CUT_D.month - r['coh'].month) >= n]
        if not den or n > 12: break
        alive = 0
        for r in den:
            cd = r['churn_date']
            if cd is None or ((cd.year - r['coh'].year) * 12 + (cd.month - r['coh'].month)) >= n: alive += 1
        o.append([n, alive, len(den), alive / len(den)]); n += 1
    surv[prod] = o
dump('survival', surv)

# ───────── 5. константы карточки ─────────
c = {}
c['SRV'] = {p: sum(1 for b in base if b['prod'] == p) for p in ('POS', 'SFA')}
c['MRR_SUBS'] = {p: sum(b['mrr'] for b in base if b['prod'] == p) for p in ('POS', 'SFA')}   # по подпискам (справочно)
# MRR на срез = нарастающий итог журнала изменений MRR (sale.order.log) по дату среза — как график «Изменение MRR» в Odoo
_lj = {r['prod']: float(r['s']) for r in db.query(f"""select {PROD_SQL} prod, sum(case when l.currency_id=1 then l.amount_signed*12000 else l.amount_signed end) s
   from sale_order_log l join sale_order so on so.id=l.order_id where l.event_date <= date '{CUT}' group by 1""")}
c['MRR'] = {p: _lj.get(p, 0.0) for p in ('POS', 'SFA')}
c['SPLIT'] = {p: {s: [sum(1 for b in base if b['prod']==p and b['st']==s), sum(b['mrr'] for b in base if b['prod']==p and b['st']==s)]
                  for s in ('3_progress','4_paused')} for p in ('POS','SFA')}
bp = {b['id']: b['prod'] for b in base}
REC = {24, 40, 43, 25, 26, 41, 42, 35, 32}
lic = collections.defaultdict(float)
for l in lines:
    if l['pid'] in REC: lic[bp[l['oid']]] += l['qty']
c['LIC'] = dict(lic)
c['NEW'] = {r['prod']: r['n'] for r in db.query(f"""select {PROD_SQL} prod, count(*) n from sale_order_log l join sale_order so on so.id=l.order_id
 where l.event_type='0_creation' and l.event_date>=date '{P1}' and l.event_date<date '{P2}' group by 1""")}
c['NEW_COH'] = {p: sum(r['n'] for r in out['cohorts'] if r['prod'] == p) for p in ('POS', 'SFA')}
ONE = {'POS': (30, 39), 'SFA': (33, 37, 35, 36, 38, 44)}
allone = list(ONE['POS']) + list(ONE['SFA'])
om = {p: {m: 0.0 for m in MON} for p in ONE}; byp = collections.defaultdict(float)
for r in db.query(f"""select to_char(aml.date,'YYYY-MM') m, aml.product_id pid,
   sum(case when am.currency_id=1 then aml.price_subtotal*12000 else aml.price_subtotal end) amt
 from account_move_line aml join account_move am on am.id=aml.move_id
 where am.move_type in ('out_invoice','out_refund') and aml.parent_state='posted' and aml.display_type='product'
   and aml.date>=date '{P1}' and aml.date<date '{P2}' and aml.product_id = any(%s) group by 1,2""", (allone,)):
    byp[r['pid']] += float(r['amt'] or 0)
    for p, pids in ONE.items():
        if r['pid'] in pids and r['m'] in om[p]: om[p][r['m']] += float(r['amt'] or 0)
c['ONCE_M'] = {p: [om[p][m] for m in MON] for p in om}; c['ONCE'] = {p: sum(om[p].values()) for p in om}
c['ONCE_BY_PID'] = {str(k): v for k, v in byp.items()}
dep = {r['dep']: float(r['amt'] or 0) for r in db.query(f"""select coalesce(hd.name->>'ru_RU', hd.name->>'en_US','(нет)') dep, sum(l.x_studio_sum) amt
 from x_custom_expenses_line_dc8d4 l left join hr_department hd on hd.id=l.x_studio_department
 where l.x_studio_exp_date>=date '{P1}' and l.x_studio_exp_date<date '{P2}' group by 1""")}
pick = lambda *keys: sum(v for k, v in dep.items() if any(x.lower() in k.lower() for x in keys))
c['DEP'] = {'sales_POS': pick('продаж POS'), 'sales_SFA': pick('продаж SFA'), 'marketing': pick('Маркетинг'),
            'onb_POS': pick('Onboarding POS'), 'onb_SFA': pick('Onboarding SFA'), 'support': pick('тех. поддержки'),
            'ka': pick('ключевыми клиент'), 'partner': pick('партнер')}
c['DEP_ALL'] = dep
ch = collections.defaultdict(int)
for r in out['events_n']:
    if r['et'] == '2_churn': ch[(r['prod'], r['m'])] = r['n']
c['CHURN_N'] = {p: {m: ch[(p, m)] for m in MON} for p in ('POS', 'SFA')}
def cycles(rw):
    """Срок жизни по когортной кривой: наблюдаемые месяцы + хвост по среднему темпу убыли последних трёх месяцев."""
    obs = sum(r[3] for r in rw[1:]); last = rw[-1][3]
    if len(rw) >= 5 and rw[-4][3]: p = (rw[-1][3] / rw[-4][3]) ** (1 / 3)
    else: p = (rw[-1][3] / rw[-2][3]) if len(rw) > 1 and rw[-2][3] else 0
    p = min(p, 0.98)
    return obs + (last * p / (1 - p) if 0 < p < 1 else 0)
c['CYC_COH'] = {p: cycles(surv[p]) for p in surv}

# ───────── 6. расходы: хостинг и затраты на привлечение ─────────
AMT = """(coalesce((select sum(l.x_studio_sum) from x_custom_expenses_line_dc8d4 l where l.x_custom_expenses_id=h.id),0)
          * case when coalesce(h.x_studio_rate,0)>0 then h.x_studio_rate else 1 end)"""
host = db.query(f"""select h.id, h.x_studio_date pay, h.x_studio_start_date d1, h.x_studio_end_date d2, coalesce(h.x_studio_rate,0) rate,
  coalesce((select sum(l.x_studio_sum) from x_custom_expenses_line_dc8d4 l where l.x_custom_expenses_id=h.id),0) lsum, h.x_studio_notes nt
 from x_custom_expenses h where h.x_studio_accountanalyticaccount=605 and h.x_studio_start_date<=date '{CUT}'
 order by h.x_studio_start_date, h.id""")
clean = lambda s: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', s or '')).strip()
dump('hosting', [{'id': r['id'], 'pay': str(r['pay']), 'd1': str(r['d1']), 'd2': str(r['d2']), 'rate': float(r['rate']),
                  'lsum': float(r['lsum']), 'note': clean(r['nt'])} for r in host])
pools = {}
for key, aid in (('AD', 601), ('CRM', 603), ('COMM', 622)):
    pools[key] = float(db.query(f"""select coalesce(sum({AMT}),0) a from x_custom_expenses h where h.x_studio_accountanalyticaccount=%s
      and h.x_studio_start_date>=date '{P1}' and h.x_studio_start_date<date '{P2}'""", (aid,))[0]['a'])
pools['MKT_CONTR'] = float(db.query(f"""select coalesce(sum({AMT}),0) a from x_custom_expenses h where h.x_studio_accountanalyticaccount=579
  and h.x_studio_start_date>=date '{P1}' and h.x_studio_start_date<date '{P2}'
  and (h.x_studio_notes ilike '%%монтаж%%' or h.x_studio_notes ilike '%%дизайнер%%')""")[0]['a'])
pools['AD_N'] = db.query(f"""select count(*) n from x_custom_expenses h where h.x_studio_accountanalyticaccount=601
  and h.x_studio_start_date>=date '{P1}' and h.x_studio_start_date<date '{P2}'""")[0]['n']
pools['COMM_N'] = db.query(f"""select count(*) n from x_custom_expenses h where h.x_studio_accountanalyticaccount=622
  and h.x_studio_start_date>=date '{P1}' and h.x_studio_start_date<date '{P2}'""")[0]['n']
c['POOLS'] = pools

# ───────── 7. штат отделов себестоимости на срез ─────────
staff = db.query(f"""select e.name, e.department_id dep, coalesce(nullif(e.job_title,''), j.name->>'ru_RU') job,
   coalesce(e.x_studio_salary_1,0) sal, e.active, e.departure_date dd, e.create_date::date cd, (d.manager_id=e.id) is_head
 from hr_employee e left join hr_job j on j.id=e.job_id left join hr_department d on d.id=e.department_id
 where e.department_id in (4,13,10,12) and e.create_date::date <= date '{CUT}'
   and (e.active or (e.departure_date is not null and e.departure_date > date '{CUT}'))""")
dump('staff', [{'name': r['name'], 'dep': r['dep'], 'job': (r['job'] or '').strip(), 'sal': float(r['sal']),
                'head': bool(r['is_head'])} for r in staff])
dump('const', c)
for k in ('SRV', 'LIC', 'MRR', 'NEW', 'NEW_COH', 'ONCE', 'CYC_COH', 'POOLS'): print(k, c[k])
print('DEP', {k: f'{v:,.0f}' for k, v in c['DEP'].items()})
print('survival', {p: [f"{x[0]}:{x[3]*100:.0f}%" for x in surv[p]] for p in surv})
tot = {p: 0.0 for p in ('POS', 'SFA')}
for r in out['moves']:
    if r['prod'] in tot: tot[r['prod']] += r['amt']
print('MRR0 (обратным счётом):', {p: int(c['MRR'][p] - tot[p]) for p in tot})
print('штат:', collections.Counter(s['dep'] for s in json.load(open(J('staff')))))
