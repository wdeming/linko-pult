# -*- coding: utf-8 -*-
"""Затраты на привлечение и каналы для карточки метрик: факт по месяцам с разметкой, каналы лидов, обучение POS."""
import json, collections
from odoo_db import OdooDB
from cls import PROD_SQL
from cfg_cut import *
db = OdooDB(); o = {}
# ФОТ по отделам помесячно (разметка отделов в строках расходов ведётся с мая 2026)
rows = db.query(f"""select to_char(l.x_studio_exp_date,'YYYY-MM') m, l.x_studio_department dep, sum(l.x_studio_sum) amt
 from x_custom_expenses_line_dc8d4 l where l.x_studio_exp_date>=date '{P1}' and l.x_studio_exp_date<date '{P2}'
 and l.x_studio_department is not null group by 1,2""")
DEP = {3: 'sales_POS', 9: 'sales_SFA', 8: 'marketing', 13: 'onb_POS', 10: 'onb_SFA', 15: 'partner'}
fm = {v: {m: 0.0 for m in MON} for v in DEP.values()}
for r in rows:
    if r['dep'] in DEP: fm[DEP[r['dep']]][r['m']] += float(r['amt'])
o['FOT_M'] = fm
AMT = """(coalesce((select sum(l.x_studio_sum) from x_custom_expenses_line_dc8d4 l where l.x_custom_expenses_id=h.id),0)
          * case when coalesce(h.x_studio_rate,0)>0 then h.x_studio_rate else 1 end)"""
pm = {}
for key, aid, extra in (('AD', 601, ''), ('CRM', 603, ''), ('COMM', 622, ''),
                        ('MKT_CONTR', 579, " and (h.x_studio_notes ilike '%%монтаж%%' or h.x_studio_notes ilike '%%дизайнер%%')")):
    d = {m: 0.0 for m in MON}
    for r in db.query(f"""select to_char(h.x_studio_start_date,'YYYY-MM') m, sum({AMT}) a from x_custom_expenses h
      where h.x_studio_accountanalyticaccount={aid} and h.x_studio_start_date>=date '{P1}' and h.x_studio_start_date<date '{P2}' {extra} group by 1"""):
        d[r['m']] = float(r['a'] or 0)
    pm[key] = d
o['POOL_M'] = pm
# каналы: подписки когорты по источнику лида и активные лиды, созданные в периоде
subs = collections.defaultdict(lambda: {'POS': 0, 'SFA': 0})
for r in db.query(f"""select coalesce(us.name,'(нет)') src, {PROD_SQL} prod, count(*) n
  from sale_order so left join crm_lead l on l.id=so.opportunity_id left join utm_source us on us.id=l.source_id
  where so.first_contract_date>=date '{P1}' and so.first_contract_date<date '{P2}'
    and exists (select 1 from sale_order_log g where g.order_id=so.id) group by 1,2"""):
    if r['prod'] in ('POS', 'SFA'): subs[r['src']][r['prod']] += r['n']
leads = {r['src']: r['n'] for r in db.query(f"""select coalesce(us.name,'(нет)') src, count(*) n
  from crm_lead l left join utm_source us on us.id=l.source_id
  where l.active and l.create_date>=date '{P1}' and l.create_date<date '{P2}' group by 1""")}
o['CH_SUBS'] = dict(subs); o['CH_LEADS'] = leads
e = db.query(f"""
 with inv as (select am.partner_id pid, aml.date d, aml.price_subtotal amt from account_move_line aml join account_move am on am.id=aml.move_id
    where aml.product_id=30 and aml.parent_state='posted' and am.move_type='out_invoice' and aml.display_type='product'
      and aml.date>=date '{P1}' and aml.date<date '{P2}'),
  fc as (select partner_id pid, min(first_contract_date) fc from sale_order where first_contract_date is not null group by 1)
 select count(distinct inv.pid) clients, sum(amt) amt,
   sum(case when fc.fc>=date '{P1}' then amt else 0 end)/sum(amt) new_share,
   avg(case when fc.fc>=date '{P1}' then (inv.d-fc.fc) end) lag
 from inv left join fc on fc.pid=inv.pid""")[0]
o['EDU'] = {k: float(v) for k, v in e.items()}
json.dump(o, open(J('mkt'), 'w'), ensure_ascii=False)
for k in ('sales_POS', 'sales_SFA', 'marketing', 'onb_POS', 'onb_SFA', 'partner'):
    print(k, [round(fm[k][m] / 1e6, 1) for m in MON])
for k in pm: print(k, [round(pm[k][m] / 1e6, 1) for m in MON])
print({k: v for k, v in o['CH_SUBS'].items() if k in ('Инстаграм', 'Фейсбук', 'Сайт', 'Партнер')}, {k: leads.get(k) for k in ('Инстаграм', 'Фейсбук', 'Сайт')})
print(o['EDU'])
