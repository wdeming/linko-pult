# -*- coding: utf-8 -*-
"""Отток: закрытые подписки с причинами (на срез cfg_cut.CUT) и помесячная база для листа «Динамика оттока»."""
import json, datetime as dt
from odoo_db import OdooDB
from cls import PROD_SQL, MRR_SQL, AMT_SQL
from cfg_cut import *
db = OdooDB()
out = {}
# закрытые подписки: сейчас в статусе «отток», причина указана, дата закрытия (последнее событие оттока в журнале, иначе end_date) в периоде
q = f"""
with x as (select so.id, so.partner_id, {PROD_SQL} prod, so.close_reason_id cr, {MRR_SQL} mrr,
   coalesce((select max(l.event_date) from sale_order_log l where l.order_id=so.id and l.event_type='2_churn'), so.end_date) cd
 from sale_order so where so.subscription_state='6_churn')
select prod, cr, count(*) n, sum(mrr) m, array_agg(distinct partner_id) partners from x
where cd >= date '{P1}' and cd <= date '{CUT}' group by 1,2 order by 1,2"""
out['closed'] = [{'prod': r['prod'], 'cr': r['cr'], 'n': r['n'], 'm': float(r['m'] or 0), 'partners': r['partners']} for r in db.query(q)]
out['reasons'] = {str(r['id']): r['ru'] for r in db.query("select id, coalesce(name->>'ru_RU', name->>'en_US') ru from sale_order_close_reason")}
# база подписок на начало каждого месяца — то же правило, что и для среза
base = {}
for m in MON:
    d0 = (dt.date.fromisoformat(m + '-01') - dt.timedelta(days=1)).isoformat()
    q = f"""
    with ev as (select order_id,
        max(case when event_type='0_creation' and event_date > date '{d0}' then 1 else 0 end) ca,
        max(case when event_type='2_churn' and event_date > date '{d0}' then 1 else 0 end) ch
      from sale_order_log group by order_id)
    select {PROD_SQL} prod, count(*) n from sale_order so left join ev on ev.order_id=so.id
    where coalesce(ev.ca,0)=0 and (so.subscription_state in ('3_progress','4_paused') or coalesce(ev.ch,0)=1) group by 1"""
    base[m] = {r['prod']: r['n'] for r in db.query(q)}
out['base_n'] = base
json.dump(out, open(J('churn'), 'w'), ensure_ascii=False)
for p in ('POS', 'SFA'):
    rs = [r for r in out['closed'] if r['prod'] == p]
    print(p, 'закрыто', sum(r['n'] for r in rs), 'с причиной', sum(r['n'] for r in rs if r['cr']), 'MRR', int(sum(r['m'] for r in rs if r['cr'])),
          '| без причины', sum(r['n'] for r in rs if not r['cr']), int(sum(r['m'] for r in rs if not r['cr'])))
print({m: base[m] for m in MON})
